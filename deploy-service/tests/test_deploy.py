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


def test_wait_for_endpoint_passes_headers_to_request():
    received = {}

    def handler(request: httpx.Request) -> httpx.Response:
        received["host"] = request.headers.get("host")
        return httpx.Response(200)

    client = httpx.Client(transport=httpx.MockTransport(handler))

    assert (
        wait_for_endpoint(
            client,
            "http://openl/REST/shop-policy",
            headers={"Host": "example.com:8080"},
            timeout=1,
            interval=0.01,
        )
        is True
    )
    assert received["host"] == "example.com:8080"
