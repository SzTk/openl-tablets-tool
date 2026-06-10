FROM python:3.12-slim

RUN pip install --no-cache-dir uv

WORKDIR /app

COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev

COPY service.py templates.py deploy.py ./

EXPOSE 8000

CMD ["uv", "run", "uvicorn", "service:app", "--host", "0.0.0.0", "--port", "8000"]
