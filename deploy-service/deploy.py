import time

import httpx


def wait_for_endpoint(
    client: httpx.Client,
    url: str,
    headers: dict[str, str] | None = None,
    timeout: float = 60.0,
    interval: float = 2.0,
) -> bool:
    """Poll `url` until it returns 200 OK or `timeout` seconds elapse."""
    deadline = time.monotonic() + timeout

    while True:
        try:
            response = client.get(url, headers=headers)
            if response.status_code == 200:
                return True
        except httpx.RequestError:
            pass

        if time.monotonic() >= deadline:
            return False
        time.sleep(interval)
