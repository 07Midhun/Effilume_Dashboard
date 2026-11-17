FROM python:3.12-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copy dependency definition
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create persistent folders
RUN mkdir -p uploads outputs

EXPOSE 8000

ENV HOST=0.0.0.0
ENV PORT=8000

CMD ["python", "run_server.py"]
