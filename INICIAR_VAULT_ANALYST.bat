@echo off
title VaultAnalyst AI - Servidor Privado Localhost
color 0A
echo ==============================================================
echo       VAULTANALYST AI - ARQUITECTURA ZERO-TRUST LOCALHOST
echo ==============================================================
echo [INFO] Iniciando aplicacion web privada en tu navegador...
echo [INFO] Los datos NO salen a internet. Procesamiento 100%% en RAM.
echo.
cd /d "%~dp0"
.\backend\.venv\Scripts\python.exe -m streamlit run backend/app_vault_analyst.py --server.headless=false --server.port=8501
pause
