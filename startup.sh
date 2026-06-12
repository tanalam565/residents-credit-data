#!/bin/bash
# Azure App Service startup script

# Use /home/data for persistent storage (survives restarts & deployments)
mkdir -p /home/data
export DB_PATH="${DB_PATH:-/home/data/reports.db}"

cd /home/site/wwwroot/backend

# Activate the virtual environment built during deployment
source /home/site/wwwroot/antenv/bin/activate

gunicorn \
  --workers 1 \
  --worker-class uvicorn.workers.UvicornWorker \
  --bind 0.0.0.0:8000 \
  --timeout 120 \
  app:app
