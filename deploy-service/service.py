import os
import shutil
from pathlib import Path

import httpx
from fastapi import FastAPI, Form, HTTPException, UploadFile

from deploy import wait_for_endpoint
from templates import derive_service_name, generate_rules_deploy_xml, generate_rules_xml

app = FastAPI(title="OpenL Tablets Deploy Service")


def get_deployment_path() -> Path:
    return Path(os.environ.get("DEPLOYMENT_PATH", "/deployment"))


def get_openl_internal_url() -> str:
    return os.environ.get("OPENL_INTERNAL_URL", "http://openl:8080")


def get_openl_public_url() -> str:
    return os.environ.get("OPENL_PUBLIC_URL", "http://localhost:9080")


def get_deploy_timeout() -> float:
    return float(os.environ.get("OPENL_DEPLOY_TIMEOUT", "60"))


def get_deploy_interval() -> float:
    return float(os.environ.get("OPENL_DEPLOY_INTERVAL", "2"))


def get_http_client() -> httpx.Client:
    return httpx.Client()


def _sanitize_name(name: str) -> str:
    if not name or name in (".", "..") or "/" in name or "\\" in name:
        raise HTTPException(status_code=400, detail=f"Invalid service name: '{name}'")
    return name


@app.get("/health")
def health() -> dict:
    return {"status": "ok"}


@app.get("/services")
def list_services() -> dict:
    deployment_path = get_deployment_path()
    if not deployment_path.exists():
        return {"services": []}
    return {
        "services": sorted(
            entry.name for entry in deployment_path.iterdir() if entry.is_dir()
        )
    }


@app.delete("/services/{name}")
def delete_service(name: str) -> dict:
    name = _sanitize_name(name)
    target = get_deployment_path() / name
    if not target.is_dir():
        raise HTTPException(status_code=404, detail=f"Service '{name}' not found")
    shutil.rmtree(target)
    return {"deleted": name}


@app.post("/deploy")
async def deploy(file: UploadFile, service_name: str | None = Form(default=None)) -> dict:
    if not file.filename or not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx files are supported")

    excel_filename = Path(file.filename).name
    name = _sanitize_name(service_name) if service_name else derive_service_name(excel_filename)

    target_dir = get_deployment_path() / name
    target_dir.mkdir(parents=True, exist_ok=True)

    excel_path = target_dir / excel_filename
    with excel_path.open("wb") as out:
        shutil.copyfileobj(file.file, out)

    (target_dir / "rules.xml").write_text(generate_rules_xml(name, excel_filename))
    (target_dir / "rules-deploy.xml").write_text(generate_rules_deploy_xml(name))

    internal_endpoint = f"{get_openl_internal_url()}/REST/{name}"
    with get_http_client() as http_client:
        ready = wait_for_endpoint(
            http_client,
            internal_endpoint,
            timeout=get_deploy_timeout(),
            interval=get_deploy_interval(),
        )

    if not ready:
        raise HTTPException(
            status_code=504,
            detail=f"Timed out waiting for OpenL Tablets to deploy '{name}'",
        )

    public_endpoint = f"{get_openl_public_url()}/REST/{name}"
    return {
        "service_name": name,
        "endpoint": public_endpoint,
        "swagger_url": f"{public_endpoint}/api-docs",
    }
