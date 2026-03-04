@echo off
chcp 65001 > nul
cls

echo ================================================================================
echo    🚀 PIPELINE DE MÍDIA PAGA - EXECUÇÃO COMPLETA
echo ================================================================================
echo.
echo Este script executa todo o pipeline de análise de mídia paga:
echo   1️⃣  Análise de Performance Meta Ads
echo   2️⃣  Análise de Performance Google Ads
echo   3️⃣  Análise de Performance HubSpot (Integração)
echo   4️⃣  Geração de Cenários Preditivos
echo.
echo ================================================================================
echo.

cd /d "%~dp0"

echo 📂 Diretório de trabalho: %CD%
echo.

REM ============================================================================
REM ETAPA 1: Meta Ads
REM ============================================================================
echo.
echo ================================================================================
echo 1️⃣  ETAPA 1/4 - Análise de Performance Meta Ads
echo ================================================================================
echo.
python scripts\1.analise_performance_meta.py
if errorlevel 1 (
    echo.
    echo ❌ ERRO na Etapa 1 - Meta Ads
    echo ℹ️  Verifique os dados em: data\meta_dataset.csv
    pause
    exit /b 1
)
echo.
echo ✅ Etapa 1 concluída com sucesso!
echo.

REM ============================================================================
REM ETAPA 2: Google Ads
REM ============================================================================
echo.
echo ================================================================================
echo 2️⃣  ETAPA 2/4 - Análise de Performance Google Ads
echo ================================================================================
echo.
python scripts\2.analise_performance_google.py
if errorlevel 1 (
    echo.
    echo ❌ ERRO na Etapa 2 - Google Ads
    echo ℹ️  Verifique os dados em: data\googleads_dataset.csv
    pause
    exit /b 1
)
echo.
echo ✅ Etapa 2 concluída com sucesso!
echo.

REM ============================================================================
REM ETAPA 3: HubSpot (Integração)
REM ============================================================================
echo.
echo ================================================================================
echo 3️⃣  ETAPA 3/4 - Análise de Performance HubSpot (Integração)
echo ================================================================================
echo.
python scripts\3.analise_performance_hubspot_FINAL_ID.py
if errorlevel 1 (
    echo.
    echo ❌ ERRO na Etapa 3 - HubSpot
    echo ℹ️  Verifique os dados em: data\hubspot_leads.csv
    pause
    exit /b 1
)
echo.
echo ✅ Etapa 3 concluída com sucesso!
echo.

REM ============================================================================
REM ETAPA 4: Cenários Preditivos
REM ============================================================================
echo.
echo ================================================================================
echo 4️⃣  ETAPA 4/4 - Geração de Cenários Preditivos
echo ================================================================================
echo.
python scripts\4.gerar_cenarios_preditivos.py
if errorlevel 1 (
    echo.
    echo ❌ ERRO na Etapa 4 - Cenários Preditivos
    echo ℹ️  Verifique se os arquivos anteriores foram gerados corretamente
    pause
    exit /b 1
)
echo.
echo ✅ Etapa 4 concluída com sucesso!
echo.

REM ============================================================================
REM CONCLUSÃO
REM ============================================================================
echo.
echo ================================================================================
echo    🎉 PIPELINE EXECUTADO COM SUCESSO!
echo ================================================================================
echo.
echo 📊 Resultados gerados em: outputs\
echo.
echo Arquivos principais criados:
echo   📄 outputs\meta_dataset_dashboard.xlsx
echo   📄 outputs\google_dashboard.xlsx
echo   📄 outputs\meta_googleads_blend_[data].xlsx
echo   📄 outputs\cenarios_preditivos_[data].xlsx
echo.
echo ================================================================================
echo.

pause
