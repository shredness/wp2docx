FROM python:3.12-slim
ENV PYTHONUNBUFFERED=1 PYTHONDONTWRITEBYTECODE=1
RUN apt-get update && apt-get install -y --no-install-recommends pandoc \
    && rm -rf /var/lib/apt/lists/* \
    && pip install --no-cache-dir "requests>=2.32,<3"
COPY wp2docx.py /app/wp2docx.py
CMD ["python", "/app/wp2docx.py"]
