# OpenL Tablets Cloud Deploy Service Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let `openl-tablets-create` / `openl-tablets-edit` users deploy their Excel rule files to a running OpenL Tablets instance on Azure and test the resulting REST API end-to-end, without needing local Docker.

**Architecture:** A FastAPI "deploy service" runs alongside an `openltablets/ws` container (via docker-compose locally, and as a two-container Azure Container Instance group in the cloud). It accepts an Excel file via `POST /deploy`, writes `rules.xml` + `rules-deploy.xml` next to the Excel file in a shared `deployment/` folder that OpenL Tablets watches (`repo-file` mode), waits for OpenL to expose the new REST endpoint, and returns the public endpoint + Swagger URL. `deploy/azure/*.sh` scripts manage the on-demand ACI lifecycle (deploy/start/stop). A new Claude Code skill, `openl-tablets-deploy`, drives the whole flow after `openl-tablets-create`/`-edit`.

**Tech Stack:** Python 3.12, FastAPI + uvicorn + httpx, pytest with `httpx.MockTransport`, `uv` for dependency management, Docker / docker-compose, Azure CLI (Container Instances + Storage/Azure Files).

---

## File Structure

```
deploy-service/
├── pyproject.toml       # uv-managed deps (fastapi, uvicorn, httpx, python-multipart; dev: pytest)
├── service.py           # FastAPI app: /health, /services, /services/{name}, /deploy
├── templates.py         # derive_service_name, generate_rules_xml, generate_rules_deploy_xml
├── deploy.py            # wait_for_endpoint polling helper
├── Dockerfile
└── tests/
    ├── test_service.py
    ├── test_templates.py
    └── test_deploy.py

docker-compose.yml        # repo root: openl + deploy-service, for local verification

deploy/azure/
├── container-group.yaml.template   # ACI multi-container group definition
├── deploy.sh             # one-time: resource group, storage, file share, ACI create
├── start.sh              # az container start (after stop)
├── stop.sh               # az container stop (cost saving)
└── README.md             # setup + variables

skills/openl-tablets-deploy/
└── SKILL.md              # Claude Code skill: deploy + test flow
```

Note: `deploy-service/main.py` is NOT used as a filename — the repo's root `.gitignore` has a blanket `main.py` rule that would silently exclude it. Use `service.py`.

---

## Task 1: Deploy Service scaffolding + `/health`

**Files:**
- Create: `deploy-service/pyproject.toml`
- Create: `deploy-service/service.py`
- Create: `deploy-service/tests/test_service.py`

- [ ] **Step 1: Create `deploy-service/pyproject.toml`**

```toml
[project]
name = "openl-deploy-service"
version = "0.1.0"
description = "Upload service that deploys OpenL Tablets Excel rule files to a running OpenL Tablets instance"
requires-python = ">=3.12"
dependencies = [
    "fastapi>=0.115.0",
    "uvicorn[standard]>=0.32.0",
    "python-multipart>=0.0.12",
    "httpx>=0.27.0",
]

[dependency-groups]
dev = [
    "pytest>=8.3.0",
]

[tool.uv]
package = false
```

- [ ] **Step 2: Write the failing test for `/health`**

Create `deploy-service/tests/test_service.py`:

```python
from fastapi.testclient import TestClient

from service import app

client = TestClient(app)


def test_health_returns_ok():
    response = client.get("/health")

    assert response.status_code == 200
    assert response.json() == {"status": "ok"}
```

- [ ] **Step 3: Run test to verify it fails**

Run:
```bash
cd deploy-service && uv sync && uv run pytest -v
```
Expected: FAIL — `ModuleNotFoundError: No module named 'service'` (or collection error), since `service.py` does not exist yet.

- [ ] **Step 4: Create `deploy-service/service.py` with the FastAPI app and `/health`**

```python
from fastapi import FastAPI

app = FastAPI(title="OpenL Tablets Deploy Service")


@app.get("/health")
def health() -> dict:
    return {"status": "ok"}
```

- [ ] **Step 5: Run test to verify it passes**

Run:
```bash
cd deploy-service && uv run pytest -v
```
Expected: `test_service.py::test_health_returns_ok PASSED`

- [ ] **Step 6: Commit**

```bash
git add deploy-service/pyproject.toml deploy-service/uv.lock deploy-service/service.py deploy-service/tests/test_service.py
git commit -m "feat(deploy-service): scaffold FastAPI app with /health endpoint"
```

---

## Task 2: `templates.py` — service name derivation + XML generation

**Files:**
- Create: `deploy-service/templates.py`
- Create: `deploy-service/tests/test_templates.py`

- [ ] **Step 1: Write the failing tests**

Create `deploy-service/tests/test_templates.py`:

```python
from templates import derive_service_name, generate_rules_deploy_xml, generate_rules_xml


def test_derive_service_name_from_pascal_case():
    assert derive_service_name("ShopPolicy.xlsx") == "shop-policy"


def test_derive_service_name_from_multi_word_pascal_case():
    assert derive_service_name("AutoPolicyCalculation.xlsx") == "auto-policy-calculation"


def test_derive_service_name_with_underscore_and_digit():
    assert derive_service_name("ShopPolicy2_draft.xlsx") == "shop-policy2-draft"


def test_generate_rules_xml_contains_service_name_and_filename():
    xml = generate_rules_xml("shop-policy", "ShopPolicy.xlsx")

    assert "<name>shop-policy</name>" in xml
    assert '<rules-root path="ShopPolicy.xlsx"/>' in xml


def test_generate_rules_deploy_xml_contains_service_name():
    xml = generate_rules_deploy_xml("shop-policy")

    assert "<serviceName>shop-policy</serviceName>" in xml
    assert "<url>shop-policy</url>" in xml
    assert "<publisher>RESTFUL</publisher>" in xml
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```bash
cd deploy-service && uv run pytest tests/test_templates.py -v
```
Expected: FAIL — `ModuleNotFoundError: No module named 'templates'`

- [ ] **Step 3: Create `deploy-service/templates.py`**

```python
import re


def derive_service_name(excel_filename: str) -> str:
    """Derive a URL-safe service name from an Excel filename.

    "ShopPolicy.xlsx" -> "shop-policy"
    "AutoPolicyCalculation.xlsx" -> "auto-policy-calculation"
    "ShopPolicy2_draft.xlsx" -> "shop-policy2-draft"
    """
    stem = excel_filename.rsplit(".", 1)[0]
    stem = re.sub(r"(?<!^)(?=[A-Z])", "-", stem)
    stem = re.sub(r"[^a-zA-Z0-9]+", "-", stem)
    return stem.strip("-").lower()


def generate_rules_xml(service_name: str, excel_filename: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<project>
  <name>{service_name}</name>
  <modules>
    <module>
      <name>{service_name}</name>
      <rules-root path="{excel_filename}"/>
    </module>
  </modules>
</project>
"""


def generate_rules_deploy_xml(service_name: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<rules-deploy>
  <isProvideRuntimeContext>false</isProvideRuntimeContext>
  <serviceName>{service_name}</serviceName>
  <url>{service_name}</url>
  <publishers>
    <publisher>RESTFUL</publisher>
  </publishers>
</rules-deploy>
"""
```

- [ ] **Step 4: Run tests to verify they pass**

Run:
```bash
cd deploy-service && uv run pytest tests/test_templates.py -v
```
Expected: all 5 tests PASSED

- [ ] **Step 5: Commit**

```bash
git add deploy-service/templates.py deploy-service/tests/test_templates.py
git commit -m "feat(deploy-service): add rules.xml/rules-deploy.xml template generation"
```

---

## Task 3: `deploy.py` — endpoint readiness polling helper

**Files:**
- Create: `deploy-service/deploy.py`
- Create: `deploy-service/tests/test_deploy.py`

- [ ] **Step 1: Write the failing tests**

Create `deploy-service/tests/test_deploy.py`:

```python
import httpx

from deploy import wait_for_endpoint


def test_wait_for_endpoint_returns_true_when_service_responds():
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(200)

    client = httpx.Client(transport=httpx.MockTransport(handler))

    assert wait_for_endpoint(client, "http://openl/REST/shop-policy", timeout=1, interval=0.01) is True


def test_wait_for_endpoint_returns_false_on_timeout():
    def handler(request: httpx.Request) -> httpx.Response:
        raise httpx.ConnectError("connection refused", request=request)

    client = httpx.Client(transport=httpx.MockTransport(handler))

    assert wait_for_endpoint(client, "http://openl/REST/shop-policy", timeout=0.05, interval=0.01) is False


def test_wait_for_endpoint_retries_until_ready():
    calls = {"count": 0}

    def handler(request: httpx.Request) -> httpx.Response:
        calls["count"] += 1
        if calls["count"] < 3:
            return httpx.Response(503)
        return httpx.Response(200)

    client = httpx.Client(transport=httpx.MockTransport(handler))

    assert wait_for_endpoint(client, "http://openl/REST/shop-policy", timeout=1, interval=0.01) is True
    assert calls["count"] == 3
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```bash
cd deploy-service && uv run pytest tests/test_deploy.py -v
```
Expected: FAIL — `ModuleNotFoundError: No module named 'deploy'`

- [ ] **Step 3: Create `deploy-service/deploy.py`**

```python
import time

import httpx


def wait_for_endpoint(
    client: httpx.Client,
    url: str,
    timeout: float = 60.0,
    interval: float = 2.0,
) -> bool:
    """Poll `url` until it returns a non-5xx response or `timeout` seconds elapse."""
    deadline = time.monotonic() + timeout

    while True:
        try:
            response = client.get(url)
            if response.status_code < 500:
                return True
        except httpx.RequestError:
            pass

        if time.monotonic() >= deadline:
            return False
        time.sleep(interval)
```

- [ ] **Step 4: Run tests to verify they pass**

Run:
```bash
cd deploy-service && uv run pytest tests/test_deploy.py -v
```
Expected: all 3 tests PASSED

- [ ] **Step 5: Commit**

```bash
git add deploy-service/deploy.py deploy-service/tests/test_deploy.py
git commit -m "feat(deploy-service): add endpoint readiness polling helper"
```

---

## Task 4: `GET /services`, `DELETE /services/{name}`, and name sanitization

**Files:**
- Modify: `deploy-service/service.py`
- Modify: `deploy-service/tests/test_service.py`

- [ ] **Step 1: Write the failing tests**

Append to `deploy-service/tests/test_service.py`:

```python
import pytest
from fastapi import HTTPException

import service
from service import _sanitize_name


def test_list_services_returns_empty_when_deployment_path_missing(tmp_path, monkeypatch):
    monkeypatch.setenv("DEPLOYMENT_PATH", str(tmp_path / "does-not-exist"))

    response = client.get("/services")

    assert response.status_code == 200
    assert response.json() == {"services": []}


def test_list_services_returns_sorted_directory_names(tmp_path, monkeypatch):
    (tmp_path / "shop-policy").mkdir()
    (tmp_path / "auto-policy").mkdir()
    monkeypatch.setenv("DEPLOYMENT_PATH", str(tmp_path))

    response = client.get("/services")

    assert response.json() == {"services": ["auto-policy", "shop-policy"]}


def test_delete_service_removes_directory(tmp_path, monkeypatch):
    (tmp_path / "shop-policy").mkdir()
    monkeypatch.setenv("DEPLOYMENT_PATH", str(tmp_path))

    response = client.delete("/services/shop-policy")

    assert response.status_code == 200
    assert response.json() == {"deleted": "shop-policy"}
    assert not (tmp_path / "shop-policy").exists()


def test_delete_service_missing_returns_404(tmp_path, monkeypatch):
    monkeypatch.setenv("DEPLOYMENT_PATH", str(tmp_path))

    response = client.delete("/services/does-not-exist")

    assert response.status_code == 404


def test_sanitize_name_accepts_valid_name():
    assert _sanitize_name("shop-policy") == "shop-policy"


@pytest.mark.parametrize("name", ["..", ".", "../etc", "a/b", "a\\b", ""])
def test_sanitize_name_rejects_invalid_names(name):
    with pytest.raises(HTTPException) as exc_info:
        _sanitize_name(name)
    assert exc_info.value.status_code == 400
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```bash
cd deploy-service && uv run pytest tests/test_service.py -v
```
Expected: FAIL — `ImportError: cannot import name '_sanitize_name' from 'service'` and 404s for `/services` (route not found).

- [ ] **Step 3: Add the endpoints and helpers to `deploy-service/service.py`**

Replace the full contents of `deploy-service/service.py` with:

```python
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run:
```bash
cd deploy-service && uv run pytest -v
```
Expected: all tests in `test_service.py`, `test_templates.py`, and `test_deploy.py` PASS.

- [ ] **Step 5: Commit**

```bash
git add deploy-service/service.py deploy-service/tests/test_service.py
git commit -m "feat(deploy-service): add /services list and delete endpoints"
```

---

## Task 5: `POST /deploy` endpoint

**Files:**
- Modify: `deploy-service/service.py`
- Modify: `deploy-service/tests/test_service.py`

- [ ] **Step 1: Write the failing tests**

Append to `deploy-service/tests/test_service.py`:

```python
import httpx


def _ok_handler(request: httpx.Request) -> httpx.Response:
    return httpx.Response(200)


def test_deploy_writes_files_and_returns_endpoint(tmp_path, monkeypatch):
    monkeypatch.setenv("DEPLOYMENT_PATH", str(tmp_path))
    monkeypatch.setenv("OPENL_INTERNAL_URL", "http://openl:8080")
    monkeypatch.setenv("OPENL_PUBLIC_URL", "http://localhost:9080")
    monkeypatch.setattr(
        service, "get_http_client", lambda: httpx.Client(transport=httpx.MockTransport(_ok_handler))
    )

    response = client.post(
        "/deploy",
        files={"file": ("ShopPolicy.xlsx", b"fake excel content", "application/octet-stream")},
    )

    assert response.status_code == 200
    assert response.json() == {
        "service_name": "shop-policy",
        "endpoint": "http://localhost:9080/REST/shop-policy",
        "swagger_url": "http://localhost:9080/REST/shop-policy/api-docs",
    }

    deployed = tmp_path / "shop-policy"
    assert (deployed / "ShopPolicy.xlsx").read_bytes() == b"fake excel content"
    assert (deployed / "rules.xml").exists()
    assert (deployed / "rules-deploy.xml").exists()


def test_deploy_with_explicit_service_name(tmp_path, monkeypatch):
    monkeypatch.setenv("DEPLOYMENT_PATH", str(tmp_path))
    monkeypatch.setattr(
        service, "get_http_client", lambda: httpx.Client(transport=httpx.MockTransport(_ok_handler))
    )

    response = client.post(
        "/deploy",
        data={"service_name": "custom-name"},
        files={"file": ("ShopPolicy.xlsx", b"data", "application/octet-stream")},
    )

    assert response.json()["service_name"] == "custom-name"
    assert (tmp_path / "custom-name").exists()


def test_deploy_rejects_non_excel_file(tmp_path, monkeypatch):
    monkeypatch.setenv("DEPLOYMENT_PATH", str(tmp_path))

    response = client.post(
        "/deploy",
        files={"file": ("notes.txt", b"data", "text/plain")},
    )

    assert response.status_code == 400


def test_deploy_rejects_invalid_service_name(tmp_path, monkeypatch):
    monkeypatch.setenv("DEPLOYMENT_PATH", str(tmp_path))

    response = client.post(
        "/deploy",
        data={"service_name": "../escape"},
        files={"file": ("ShopPolicy.xlsx", b"data", "application/octet-stream")},
    )

    assert response.status_code == 400


def test_deploy_times_out_when_openl_never_ready(tmp_path, monkeypatch):
    monkeypatch.setenv("DEPLOYMENT_PATH", str(tmp_path))
    monkeypatch.setenv("OPENL_DEPLOY_TIMEOUT", "0.05")
    monkeypatch.setenv("OPENL_DEPLOY_INTERVAL", "0.01")

    def refused(request: httpx.Request) -> httpx.Response:
        raise httpx.ConnectError("refused", request=request)

    monkeypatch.setattr(
        service, "get_http_client", lambda: httpx.Client(transport=httpx.MockTransport(refused))
    )

    response = client.post(
        "/deploy",
        files={"file": ("ShopPolicy.xlsx", b"data", "application/octet-stream")},
    )

    assert response.status_code == 504
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```bash
cd deploy-service && uv run pytest tests/test_service.py -v
```
Expected: FAIL — `404 Not Found` for `/deploy` (route doesn't exist yet) and `AttributeError` for `service.get_http_client`.

- [ ] **Step 3: Add the `/deploy` endpoint to `deploy-service/service.py`**

Replace the full contents of `deploy-service/service.py` with:

```python
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run:
```bash
cd deploy-service && uv run pytest -v
```
Expected: all tests PASS (test_service.py, test_templates.py, test_deploy.py).

- [ ] **Step 5: Commit**

```bash
git add deploy-service/service.py deploy-service/tests/test_service.py
git commit -m "feat(deploy-service): add POST /deploy endpoint"
```

---

## Task 6: Dockerfile for the deploy service

**Files:**
- Create: `deploy-service/Dockerfile`

- [ ] **Step 1: Write `deploy-service/Dockerfile`**

```dockerfile
FROM python:3.12-slim

RUN pip install --no-cache-dir uv

WORKDIR /app

COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev

COPY service.py templates.py deploy.py ./

EXPOSE 8000

CMD ["uv", "run", "uvicorn", "service:app", "--host", "0.0.0.0", "--port", "8000"]
```

- [ ] **Step 2: Build the image to verify it works**

Run:
```bash
cd deploy-service && docker build -t openl-deploy-service .
```
Expected: image builds successfully (`Successfully tagged openl-deploy-service:latest`).

If Docker is not available in this environment, skip this verification and note it for the user to run manually.

- [ ] **Step 3: Commit**

```bash
git add deploy-service/Dockerfile
git commit -m "feat(deploy-service): add Dockerfile"
```

---

## Task 7: docker-compose for local end-to-end verification

**Files:**
- Create: `docker-compose.yml` (repo root)

- [ ] **Step 1: Write `docker-compose.yml`**

```yaml
services:
  openl:
    image: openltablets/ws
    ports:
      - "9080:8080"
    environment:
      - production-repository.factory=repo-file
      - production-repository.uri=/tmp/openl
    volumes:
      - deployment:/tmp/openl/deployment

  deploy-service:
    build: ./deploy-service
    ports:
      - "8000:8000"
    volumes:
      - deployment:/deployment
    environment:
      - DEPLOYMENT_PATH=/deployment
      - OPENL_INTERNAL_URL=http://openl:8080
      - OPENL_PUBLIC_URL=http://localhost:9080
    depends_on:
      - openl

volumes:
  deployment:
```

- [ ] **Step 2: Verify end-to-end locally**

Run:
```bash
docker compose up -d
```

Wait for OpenL Tablets to finish starting (check logs):
```bash
docker compose logs -f openl
```
Expected: log line indicating the webapp has started (e.g. `Server startup in ... ms`). Press Ctrl+C once seen.

Deploy the sample `ShopPolicy.xlsx`:
```bash
curl -X POST http://localhost:8000/deploy -F "file=@ShopPolicy.xlsx"
```
Expected JSON response:
```json
{"service_name": "shop-policy", "endpoint": "http://localhost:9080/REST/shop-policy", "swagger_url": "http://localhost:9080/REST/shop-policy/api-docs"}
```

Confirm OpenL Tablets picked up the deployment:
```bash
curl http://localhost:9080/REST/shop-policy/api-docs
```
Expected: an OpenAPI/Swagger JSON document for the `shop-policy` service (HTTP 200).

Tear down:
```bash
docker compose down -v
```

If Docker is not available in this environment, document this as a manual verification step for the user (this is the primary "does the hot-reload assumption from the design spec hold" check called out in `docs/superpowers/specs/2026-06-09-openl-cloud-deploy-design.md`).

- [ ] **Step 3: Commit**

```bash
git add docker-compose.yml
git commit -m "feat: add docker-compose for local OpenL + deploy-service verification"
```

---

## Task 8: Azure Container Instances deployment scripts

**Files:**
- Create: `deploy/azure/container-group.yaml.template`
- Create: `deploy/azure/deploy.sh`
- Create: `deploy/azure/start.sh`
- Create: `deploy/azure/stop.sh`
- Create: `deploy/azure/README.md`

- [ ] **Step 1: Write `deploy/azure/container-group.yaml.template`**

```yaml
apiVersion: 2021-09-01
location: ${LOCATION}
name: ${CONTAINER_GROUP}
properties:
  osType: Linux
  ipAddress:
    type: Public
    ports:
      - protocol: tcp
        port: 8080
      - protocol: tcp
        port: 8000
  containers:
    - name: openl
      properties:
        image: openltablets/ws
        ports:
          - port: 8080
        environmentVariables:
          - name: production-repository.factory
            value: repo-file
          - name: production-repository.uri
            value: /tmp/openl
        resources:
          requests:
            cpu: 1
            memoryInGB: 1.5
        volumeMounts:
          - name: deployment
            mountPath: /tmp/openl/deployment
    - name: deploy-service
      properties:
        image: ${DEPLOY_SERVICE_IMAGE}
        ports:
          - port: 8000
        environmentVariables:
          - name: DEPLOYMENT_PATH
            value: /deployment
          - name: OPENL_INTERNAL_URL
            value: http://localhost:8080
          - name: OPENL_PUBLIC_URL
            value: ${OPENL_PUBLIC_URL}
        resources:
          requests:
            cpu: 0.5
            memoryInGB: 0.5
        volumeMounts:
          - name: deployment
            mountPath: /deployment
  volumes:
    - name: deployment
      azureFile:
        shareName: ${SHARE_NAME}
        storageAccountName: ${STORAGE_ACCOUNT}
        storageAccountKey: ${STORAGE_KEY}
```

Note: within an ACI container group, all containers share one network namespace, so `deploy-service` reaches OpenL via `http://localhost:8080`. Both ports (`8080` for OpenL, `8000` for deploy-service) are exposed on the group's single public IP — there is no 9080→8080 remapping like in docker-compose, so `OPENL_PUBLIC_URL` must use port `8080`.

- [ ] **Step 2: Write `deploy/azure/deploy.sh`**

```bash
#!/usr/bin/env bash
set -euo pipefail

# Required: build and push the deploy-service image first, e.g.
#   docker build -t <registry>/openl-deploy-service:latest deploy-service
#   docker push <registry>/openl-deploy-service:latest
: "${DEPLOY_SERVICE_IMAGE:?Set DEPLOY_SERVICE_IMAGE to the pushed deploy-service image}"

RESOURCE_GROUP="${RESOURCE_GROUP:-openl-demo-rg}"
LOCATION="${LOCATION:-japaneast}"
STORAGE_ACCOUNT="${STORAGE_ACCOUNT:-openldemostorage}"
SHARE_NAME="${SHARE_NAME:-openl-deployment}"
CONTAINER_GROUP="${CONTAINER_GROUP:-openl-demo}"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

az group create --name "$RESOURCE_GROUP" --location "$LOCATION" --output none

az storage account create \
  --name "$STORAGE_ACCOUNT" \
  --resource-group "$RESOURCE_GROUP" \
  --location "$LOCATION" \
  --sku Standard_LRS \
  --output none

STORAGE_KEY=$(az storage account keys list \
  --account-name "$STORAGE_ACCOUNT" \
  --resource-group "$RESOURCE_GROUP" \
  --query '[0].value' -o tsv)

az storage share create \
  --name "$SHARE_NAME" \
  --account-name "$STORAGE_ACCOUNT" \
  --account-key "$STORAGE_KEY" \
  --output none

OPENL_PUBLIC_URL="${OPENL_PUBLIC_URL:-}"
if [[ -z "$OPENL_PUBLIC_URL" ]]; then
  # Placeholder; ACI assigns the public IP only after creation. Re-run
  # start.sh to print the IP, then update the running group if needed.
  OPENL_PUBLIC_URL="http://PENDING:8080"
fi

export LOCATION CONTAINER_GROUP DEPLOY_SERVICE_IMAGE OPENL_PUBLIC_URL SHARE_NAME STORAGE_ACCOUNT STORAGE_KEY

envsubst < "$SCRIPT_DIR/container-group.yaml.template" > "$SCRIPT_DIR/container-group.yaml"

az container create \
  --resource-group "$RESOURCE_GROUP" \
  --file "$SCRIPT_DIR/container-group.yaml" \
  --output none

az container show \
  --name "$CONTAINER_GROUP" \
  --resource-group "$RESOURCE_GROUP" \
  --query "{state:instanceView.state, ip:ipAddress.ip}" -o table
```

- [ ] **Step 3: Write `deploy/azure/start.sh`**

```bash
#!/usr/bin/env bash
set -euo pipefail

RESOURCE_GROUP="${RESOURCE_GROUP:-openl-demo-rg}"
CONTAINER_GROUP="${CONTAINER_GROUP:-openl-demo}"

az container start --name "$CONTAINER_GROUP" --resource-group "$RESOURCE_GROUP"

az container show \
  --name "$CONTAINER_GROUP" \
  --resource-group "$RESOURCE_GROUP" \
  --query "{state:instanceView.state, ip:ipAddress.ip}" -o table
```

- [ ] **Step 4: Write `deploy/azure/stop.sh`**

```bash
#!/usr/bin/env bash
set -euo pipefail

RESOURCE_GROUP="${RESOURCE_GROUP:-openl-demo-rg}"
CONTAINER_GROUP="${CONTAINER_GROUP:-openl-demo}"

az container stop --name "$CONTAINER_GROUP" --resource-group "$RESOURCE_GROUP"
```

- [ ] **Step 5: Make scripts executable**

Run:
```bash
chmod +x deploy/azure/deploy.sh deploy/azure/start.sh deploy/azure/stop.sh
```

- [ ] **Step 6: Write `deploy/azure/README.md`**

```markdown
# Azure へのデプロイ (Container Instances)

OpenL Tablets (`openltablets/ws`) と deploy-service を同一の Azure Container Instance (ACI) コンテナグループにデプロイする。

## 前提

- Azure CLI (`az`) がインストール済みで `az login` 済みであること
- deploy-service イメージをビルドし、Azure 上から取得可能なレジストリに push 済みであること

```bash
docker build -t <registry>/openl-deploy-service:latest ../../deploy-service
docker push <registry>/openl-deploy-service:latest
```

## 初回デプロイ

```bash
export DEPLOY_SERVICE_IMAGE=<registry>/openl-deploy-service:latest
export RESOURCE_GROUP=openl-demo-rg     # 任意
export LOCATION=japaneast               # 任意
./deploy.sh
```

実行後に表示される IP アドレスを使い、`OPENL_PUBLIC_URL` を `http://<ip>:8080` として
deploy-service の環境変数を更新する場合は、コンテナグループを再作成するか
`az container create` を再実行する（ACI は稼働中コンテナの環境変数を直接更新できない）。

## 起動 / 停止（コスト節約）

```bash
./start.sh   # 再起動。IP アドレスが表示される
./stop.sh    # 停止（課金を止める）
```

## エンドポイント

- OpenL Tablets REST API: `http://<ip>:8080/REST/<service-name>`
- Deploy Service: `http://<ip>:8000`
```

- [ ] **Step 7: Commit**

```bash
git add deploy/azure/
git commit -m "feat: add Azure Container Instances deployment scripts"
```

---

## Task 9: `openl-tablets-deploy` Claude Code skill

**Files:**
- Create: `skills/openl-tablets-deploy/SKILL.md`

- [ ] **Step 1: Write `skills/openl-tablets-deploy/SKILL.md`**

```markdown
---
name: openl-tablets-deploy
description: openl-tablets-create / openl-tablets-edit で作成・編集した Excel を Azure 上の OpenL Tablets インスタンスにデプロイし、REST API として動作確認する。
license: MIT
metadata:
  argument-hint: "<excel_file_path>"
allowed-tools: Bash Read
---

# OpenL Tablets クラウドデプロイ・スキル

作成・編集した Excel ルールを Azure 上の OpenL Tablets (`openltablets/ws`) にデプロイし、
REST API として正しく動作することを確認する。

## 前提

deploy-service と OpenL Tablets が Azure Container Instances にデプロイ済みであること
（`deploy/azure/README.md` 参照）。

## ステップ

### Step 1: 設定の確認

`~/.config/openl-tablets-tool/deploy.env` を確認する。

存在しない場合、ユーザーに以下を確認して保存する:

```bash
mkdir -p ~/.config/openl-tablets-tool
cat > ~/.config/openl-tablets-tool/deploy.env << 'EOF'
OPENL_DEPLOY_SERVICE_URL=http://<azure-ip>:8000
RESOURCE_GROUP=openl-demo-rg
CONTAINER_GROUP=openl-demo
EOF
```

- `OPENL_DEPLOY_SERVICE_URL`: deploy-service のベース URL
- `RESOURCE_GROUP` / `CONTAINER_GROUP`: `deploy/azure/start.sh` 用（デフォルト値で良ければそのまま）

### Step 2: サーバー起動確認

```bash
source ~/.config/openl-tablets-tool/deploy.env
curl -sf "$OPENL_DEPLOY_SERVICE_URL/health"
```

応答がない場合、ACI を起動する:

```bash
RESOURCE_GROUP="$RESOURCE_GROUP" CONTAINER_GROUP="$CONTAINER_GROUP" \
  ./deploy/azure/start.sh
```

`/health` が 200 を返すまで、5 秒間隔で最大 120 秒リトライする。
それでも応答しない場合はユーザーにエラーを報告して停止する。

### Step 3: 対象ファイルとサービス名の確認

引数 `$ARGUMENTS` が指定されていればそれを対象の Excel ファイルとする。
指定がなければ、カレントディレクトリの `.xlsx` を一覧してユーザーに選択を求める。

ファイル名からサービス名を提案する（例: `ShopPolicy.xlsx` → `shop-policy`）。
ユーザーに確認・変更の機会を与える。

### Step 4: デプロイ

```bash
curl -X POST "$OPENL_DEPLOY_SERVICE_URL/deploy" \
  -F "file=@<対象Excelの絶対パス>" \
  -F "service_name=<サービス名>"
```

レスポンスの `endpoint` と `swagger_url` を保持する。
エラー（400 / 504 など）が返った場合は内容をそのままユーザーに表示する。

### Step 5: API テスト

```bash
curl "<swagger_url>"
```

返ってきた OpenAPI 定義から、利用可能なメソッドとパラメータ型を確認する。
代表的な 1 メソッドについて、パラメータ型に合うサンプル値を組み立てて呼び出す:

```bash
curl -X POST "<endpoint>/<MethodName>" \
  -H "Content-Type: application/json" \
  -d '<サンプルJSON>'
```

レスポンスをユーザーに表示する。

### Step 6: 完了報告

```
✅ デプロイ完了

サービス名: <service_name>
エンドポイント: <endpoint>
Swagger: <swagger_url>

テスト結果:
<curl の出力>

不要になったら停止できます:
RESOURCE_GROUP=<RESOURCE_GROUP> CONTAINER_GROUP=<CONTAINER_GROUP> ./deploy/azure/stop.sh
```

## 注意事項

- `deploy.env` に保存した URL ・リソースグループ名は、ユーザー固有の Azure 環境を前提とする。
  他のユーザーが利用する場合は各自で設定する。
- Excel ファイルは `.xlsx` のみ対応。`.xls` 等は事前に変換が必要。
```

- [ ] **Step 2: Commit**

```bash
git add skills/openl-tablets-deploy/
git commit -m "feat: add openl-tablets-deploy skill for Azure E2E deployment"
```

---

## Task 10: End-to-end verification on Azure

This task is manual (requires an Azure subscription and `az login`) and is the final
acceptance check for the design spec's "エンドツーエンド" goal.

- [ ] **Step 1: Build and push the deploy-service image**

```bash
docker build -t <registry>/openl-deploy-service:latest deploy-service
docker push <registry>/openl-deploy-service:latest
```

- [ ] **Step 2: Deploy to Azure**

```bash
export DEPLOY_SERVICE_IMAGE=<registry>/openl-deploy-service:latest
./deploy/azure/deploy.sh
```

Note the printed IP address. Update `OPENL_PUBLIC_URL` to `http://<ip>:8080` if it was a
placeholder, and re-run `deploy.sh` if needed (see `deploy/azure/README.md`).

- [ ] **Step 3: Configure and run the skill**

Create `~/.config/openl-tablets-tool/deploy.env` with `OPENL_DEPLOY_SERVICE_URL=http://<ip>:8000`.

In Claude Code, invoke the `openl-tablets-deploy` skill against `ShopPolicy.xlsx` and confirm:
1. Health check passes (or ACI starts successfully)
2. `/deploy` returns `service_name`, `endpoint`, `swagger_url`
3. The test curl call against `<endpoint>/<MethodName>` returns a valid rule result (not an error)

- [ ] **Step 4: Stop the ACI to avoid ongoing charges**

```bash
RESOURCE_GROUP=openl-demo-rg CONTAINER_GROUP=openl-demo ./deploy/azure/stop.sh
```

- [ ] **Step 5: Confirm rules persist after restart**

```bash
RESOURCE_GROUP=openl-demo-rg CONTAINER_GROUP=openl-demo ./deploy/azure/start.sh
curl http://<ip>:8000/services
```
Expected: previously deployed `service_name` (e.g. `shop-policy`) is still listed, confirming
Azure Files persistence.
