from fastapi import FastAPI

app = FastAPI(title="OpenL Tablets Deploy Service")


@app.get("/health")
def health() -> dict:
    return {"status": "ok"}
