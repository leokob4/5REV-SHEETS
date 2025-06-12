@echo off
title PLM ERP - Web Server
echo 🌐 Starting FastAPI server on http://127.0.0.1:8000
python -m uvicorn client.main:app --reload
pause
