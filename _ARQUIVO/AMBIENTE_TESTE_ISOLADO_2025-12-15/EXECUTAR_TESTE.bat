@echo off
chcp 65001 > nul
echo.
echo =========================================================
echo      ROBO DE LEAD SCORING - PROVA DE FOGO
echo =========================================================
echo.
echo [1] Verificando ambiente...
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado. Instale o Python ou chame a TI.
    pause
    exit
)

echo [2] Iniciando analise...
python Scripts\3.Score_External_File.py

echo.
echo =========================================================
echo      PROCESSO FINALIZADO - Verifique a pasta Outputs
echo =========================================================
echo.
pause