@echo off
chcp 65001 > nul
cls

echo ================================================================================
echo    🎈 ANÁLISE DE FUNIL - RED BALLOON
echo ================================================================================
echo.
echo Este script gera o relatório completo de performance do funil Red Balloon
echo.
echo ================================================================================
echo.

cd /d "%~dp0"

echo 📂 Diretório de trabalho: %CD%
echo.

REM ============================================================================
REM ETAPA 1: Preparar Dados
REM ============================================================================
echo.
echo ================================================================================
echo 1️⃣  ETAPA 1/2 - Preparação de Dados
echo ================================================================================
echo.
echo Filtrando leads da Red Balloon do HubSpot completo...
echo.
python scripts\0.preparar_dados.py
if errorlevel 1 (
    echo.
    echo ❌ ERRO na preparação de dados
    echo ℹ️  Verifique se a base do HubSpot está atualizada
    echo.
    echo 💡 Execute:
    echo    python ..\\_SCRIPTS_COMPARTILHADOS\\sincronizar_bases.py
    echo.
    pause
    exit /b 1
)
echo.
echo ✅ Dados preparados com sucesso!
echo.

REM ============================================================================
REM ETAPA 2: Gerar Análise
REM ============================================================================
echo.
echo ================================================================================
echo 2️⃣  ETAPA 2/2 - Análise de Funil
echo ================================================================================
echo.
python scripts\1.analise_funil_redballoon.py
if errorlevel 1 (
    echo.
    echo ❌ ERRO na análise de funil
    pause
    exit /b 1
)
echo.
echo ✅ Análise concluída com sucesso!
echo.

REM ============================================================================
REM CONCLUSÃO
REM ============================================================================
echo.
echo ================================================================================
echo    🎉 RELATÓRIO GERADO COM SUCESSO!
echo ================================================================================
echo.
echo 📊 Resultado: outputs\funil_redballoon_[data].xlsx
echo.
echo 📑 Abas do relatório:
echo    1️⃣  Resumo Por Ciclo - Comparativo ciclo vs ciclo (Out-Mar)
echo    2️⃣  Status Atual - Distribuição dos leads
echo    3️⃣  Performance Unidade - Análise por escola
echo    4️⃣  Performance Fonte - Análise por origem
echo    5️⃣  Evolução Mensal - Tendências
echo    6️⃣  Matriz Ciclo x Unidade
echo    7️⃣  Matriz Ciclo x Fonte
echo    8️⃣  Em Qualificação - Leads ativos
echo.
echo 📅 Ciclos analisados:
echo    Ciclo 22.1: Out/2021 - Mar/2022
echo    Ciclo 23.1: Out/2022 - Mar/2023
echo    Ciclo 24.1: Out/2023 - Mar/2024
echo    Ciclo 25.1: Out/2024 - Mar/2025
echo    Ciclo 26.1: Out/2025 - Mar/2026
echo.
echo ⚠️  OBS: Leads de Abril a Setembro NÃO são incluídos
echo.
echo ================================================================================
echo.

pause
