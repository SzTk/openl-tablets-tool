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
