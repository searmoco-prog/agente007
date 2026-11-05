@echo off
chcp 65001 >nul
title VaultAnalyst AI - Servidor Privado Localhost
color 0A
echo ==============================================================
echo       VAULTANALYST AI - ARQUITECTURA ZERO-TRUST LOCALHOST
echo ==============================================================
echo [INFO] Iniciando aplicacion web privada en tu navegador...
echo [INFO] Los datos NO salen a internet. Procesamiento 100%% en RAM.
echo.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
set STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

:: Liberar el puerto 8501 si quedó una sesión previa abierta
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8501') do taskkill /f /pid %%a >nul 2>&1

start http://localhost:8501
.\backend\.venv\Scripts\python.exe -m streamlit run backend/app_vault_analyst.py --server.headless=true --server.port=8501 --browser.gatherUsageStats=false
pause
