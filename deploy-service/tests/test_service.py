import httpx
import pytest
from fastapi import HTTPException
from fastapi.testclient import TestClient

import service
from service import _sanitize_name, app

client = TestClient(app)


def test_health_returns_ok():
    response = client.get("/health")

    assert response.status_code == 200
    assert response.json() == {"status": "ok"}


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
        "endpoint": "http://localhost:9080/shop-policy",
        "swagger_url": "http://localhost:9080/shop-policy/openapi.json",
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
