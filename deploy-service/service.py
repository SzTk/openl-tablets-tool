import os
import shutil
from pathlib import Path

from fastapi import FastAPI, HTTPException

app = FastAPI(title="OpenL Tablets Deploy Service")


def get_deployment_path() -> Path:
    return Path(os.environ.get("DEPLOYMENT_PATH", "/deployment"))


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
