#!/bin/bash
# Azure App Service startup script

# Use /home/data for persistent Excel storage (survives restarts & deployments)
mkdir -p /home/data
export EXCEL_PATH="${EXCEL_PATH:-/home/data/reports.xlsx}"

cd /home/site/wwwroot

gunicorn \
  --workers 4 \
  --worker-class uvicorn.workers.UvicornWorker \
  --bind 0.0.0.0:8000 \
  --timeout 120 \
  backend.app:app
