FROM python:3.11-slim

WORKDIR /app

# Upgrade pip FIRST
RUN pip install --upgrade pip --no-cache-dir

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends curl && rm -rf /var/lib/apt/lists/*

# Copy requirements
COPY requirements_minimal.txt requirements.txt

# Install packages one by one to catch errors
RUN echo "Installing google-cloud-bigquery..." && pip install --no-cache-dir google-cloud-bigquery==3.14.1
RUN echo "Installing google-cloud-storage..." && pip install --no-cache-dir google-cloud-storage==2.10.0
RUN echo "Installing FastAPI..." && pip install --no-cache-dir fastapi==0.109.0 uvicorn[standard]==0.27.0
RUN echo "Installing openpyxl..." && pip install --no-cache-dir openpyxl==3.11.0
RUN echo "Installing pydantic..." && pip install --no-cache-dir pydantic==2.5.3
RUN echo "Installing mcp..." && pip install --no-cache-dir mcp==1.0.0

# Verify imports work BEFORE copying app code
RUN python3 -c "import google.cloud.bigquery; print('✅ BigQuery')"
RUN python3 -c "import google.cloud.storage; print('✅ Storage')"
RUN python3 -c "import fastapi; print('✅ FastAPI')"

# Copy app
COPY screportserver.py app.py

# Health check
HEALTHCHECK --interval=10s --timeout=5s --start-period=60s --retries=3 \
    CMD curl -f http://localhost:8080/health || exit 1

EXPOSE 8080

CMD ["python", "-u", "app.py"]
