@echo off
title Helio AI - Console (NAO FECHE ESTA JANELA)
echo ========================================================
echo      INICIANDO O HELIO - INTELIGENCIA ARTIFICIAL
echo ========================================================
echo.
echo [COMO COMPARTILHAR]
echo Se o link "Network URL" nao aparecer automaticamente abaixo,
echo use um dos IPs listados a seguir para acessar de outro PC:
echo.
ipconfig | findstr "IPv4"
echo.
echo DICA: O link sera http://SEU_NUMERO_IP:8501
echo.

cd /d "%~dp0"

:: Tenta rodar o Streamlit chamando o modulo do Python (mais robusto)
python -m streamlit run Scritps\Helio_App.py --server.address=0.0.0.0

:: Se o comando acima falhar (fechar sozinho), mostra erro
if %errorlevel% neq 0 (
    echo.
    echo ========================================================
    echo [ERRO] O sistema nao conseguiu iniciar.
    echo Causa provavel: O 'streamlit' nao esta instalado.
    echo SOLUCAO: Abra o terminal e digite: pip install streamlit
    echo ========================================================
    pause
)