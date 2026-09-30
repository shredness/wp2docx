FROM python:3.12-slim
RUN apt-get update && apt-get install -y --no-install-recommends pandoc \
    && rm -rf /var/lib/apt/lists/* \
    && pip install --no-cache-dir requests
COPY wp2docx.py /app/wp2docx.py
CMD ["python", "/app/wp2docx.py"]
