FROM python:3.9-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    VPA_OUTPUT_FILE=/output/vpa_recommendations.xlsx \
    VPA_LOG_LEVEL=INFO \
    VPA_LIMIT_MULTIPLIER=1.2

RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates curl \
    && curl -fsSL -o /usr/local/bin/kubectl \
        "https://dl.k8s.io/release/v1.30.4/bin/linux/amd64/kubectl" \
    && chmod +x /usr/local/bin/kubectl \
    && apt-get purge -y curl \
    && apt-get autoremove -y \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY vpa_analyzer.py .

RUN mkdir -p /output

CMD ["python", "vpa_analyzer.py"]
