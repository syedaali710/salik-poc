FROM python:3.11-slim

WORKDIR /app

COPY requirements-deploy.txt .
RUN pip install --no-cache-dir -r requirements-deploy.txt

COPY app.py chat.py heygen.py planner.py pptx_builder.py ./
COPY assets ./assets
COPY static ./static
COPY data/dashboard_data.json ./data/dashboard_data.json

ENV PORT=8000
EXPOSE 8000

CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT}"]
