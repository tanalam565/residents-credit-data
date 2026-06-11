#!/bin/bash
# Azure App Service startup script

# Use /home/data for persistent Excel storage (survives restarts & deployments)
mkdir -p /home/data
export EXCEL_PATH="${EXCEL_PATH:-/home/data/reports.xlsx}"

cd /home/site/wwwroot/backend

# Activate the virtual environment built during deployment
source /home/site/wwwroot/antenv/bin/activate

gunicorn \
  --workers 4 \
  --worker-class uvicorn.workers.UvicornWorker \
  --bind 0.0.0.0:8000 \
  --timeout 120 \
  app:app
