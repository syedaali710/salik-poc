# syntax=docker/dockerfile:1
FROM python:3.11-slim

# faster-whisper / av need ffmpeg at runtime; prebuilt av wheels target py3.11
RUN apt-get update && apt-get install -y --no-install-recommends ffmpeg \
    && rm -rf /var/lib/apt/lists/*

COPY --from=ghcr.io/astral-sh/uv:0.9.15 /uv /uvx /bin/

WORKDIR /app

ENV UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy \
    PATH="/app/.venv/bin:$PATH" \
    PORT=8000

COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev --no-install-project

COPY app.py chat.py heygen.py planner.py pptx_builder.py schemas.py ./
COPY assets ./assets
COPY static ./static
COPY data/dashboard_data.json ./data/dashboard_data.json

EXPOSE 8000
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT}"]
