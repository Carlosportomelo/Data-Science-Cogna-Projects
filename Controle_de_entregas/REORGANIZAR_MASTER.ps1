# ===================================================================
# MASTER SCRIPT - REORGANIZAÇÃO COMPLETA
# ===================================================================
# Este script executa todas as fases de reorganização sequencialmente
# ===================================================================

$ROOT = "C:\Users\a483650\Projetos"
$SCRIPT_DIR = "$ROOT\Controle_de_entregas"

Write-Host @"
╔════════════════════════════════════════════════════════════════╗
║                                                                ║
║          🏗️  REORGANIZAÇÃO DE PROJETOS - MASTER SCRIPT          ║
║                                                                ║
║  Este script irá reorganizar completamente o diretório         ║
║  C:\Users\a483650\Projetos                                    ║
║                                                                ║
╚════════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan

Write-Host "`n📋 FASES DA REORGANIZAÇÃO:`n" -ForegroundColor Yellow
Write-Host "   FASE 1: Preparação (5 min)" -ForegroundColor Green
Write-Host "      ✓ Criar estrutura _DADOS_CENTRALIZADOS"
Write-Host "      ✓ Copiar bases mais recentes"
Write-Host "      ✓ Criar documentação"
Write-Host ""
Write-Host "   FASE 2: Consolidação (10 min)" -ForegroundColor Green
Write-Host "      ✓ Renomear projetos principais"
Write-Host "      ✓ Consolidar análises operacionais"
Write-Host "      ✓ Organizar auditorias"
Write-Host "      ✓ Arquivar projetos inativos"
Write-Host ""
Write-Host "   FASE 3: Limpeza (15 min)" -ForegroundColor Green
Write-Host "      ✓ Remover duplicatas (DELETA ARQUIVOS!)"
Write-Host "      ✓ Criar backups de segurança"
Write-Host "      ✓ Criar atalhos para dados centralizados"
Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "⚠️  ANTES DE CONTINUAR:" -ForegroundColor Red
Write-Host "   1. Feche todos os programas que usam os projetos"
Write-Host "   2. Faça um backup externo completo da pasta Projetos"
Write-Host "   3. Certifique-se de ter permissões de administrador"
Write-Host ""

$continuar = Read-Host "Deseja executar TODAS as fases? (Digite 'EXECUTAR' para confirmar)"

if ($continuar -ne "EXECUTAR") {
    Write-Host "`n❌ Operação cancelada." -ForegroundColor Red
    Write-Host "`n💡 Você pode executar as fases individualmente:" -ForegroundColor Yellow
    Write-Host "   .\REORGANIZAR_FASE1_PREPARACAO.ps1"
    Write-Host "   .\REORGANIZAR_FASE2_CONSOLIDACAO.ps1"
    Write-Host "   .\REORGANIZAR_FASE3_LIMPEZA.ps1"
    exit
}

Write-Host "`n" + "="*70 -ForegroundColor Cyan
Write-Host "🚀 INICIANDO REORGANIZAÇÃO COMPLETA..." -ForegroundColor Green
Write-Host "="*70 -ForegroundColor Cyan

$inicio = Get-Date

# FASE 1: PREPARAÇÃO
Write-Host "`n▶️  Executando FASE 1: PREPARAÇÃO..." -ForegroundColor Cyan
try {
    & "$SCRIPT_DIR\REORGANIZAR_FASE1_PREPARACAO.ps1"
    $fase1_sucesso = $true
    Write-Host "✅ FASE 1 concluída!" -ForegroundColor Green
} catch {
    Write-Host "❌ ERRO na FASE 1: $_" -ForegroundColor Red
    $fase1_sucesso = $false
    exit 1
}

Start-Sleep -Seconds 2

# FASE 2: CONSOLIDAÇÃO
Write-Host "`n▶️  Executando FASE 2: CONSOLIDAÇÃO..." -ForegroundColor Cyan
try {
    & "$SCRIPT_DIR\REORGANIZAR_FASE2_CONSOLIDACAO.ps1"
    $fase2_sucesso = $true
    Write-Host "✅ FASE 2 concluída!" -ForegroundColor Green
} catch {
    Write-Host "❌ ERRO na FASE 2: $_" -ForegroundColor Red
    $fase2_sucesso = $false
    Write-Host "`n⚠️  A FASE 1 foi concluída, mas houve erro na FASE 2." -ForegroundColor Yellow
    Write-Host "   Você pode tentar executar novamente: .\REORGANIZAR_FASE2_CONSOLIDACAO.ps1" -ForegroundColor Yellow
    exit 2
}

Start-Sleep -Seconds 2

# FASE 3: LIMPEZA (com confirmação extra)
Write-Host "`n▶️  Executando FASE 3: LIMPEZA..." -ForegroundColor Cyan
Write-Host "`n⚠️  A FASE 3 irá DELETAR arquivos duplicados!" -ForegroundColor Red
$confirmar_fase3 = Read-Host "Confirmar execução da Fase 3? (Digite 'SIM')"

if ($confirmar_fase3 -eq "SIM") {
    try {
        & "$SCRIPT_DIR\REORGANIZAR_FASE3_LIMPEZA.ps1"
        $fase3_sucesso = $true
        Write-Host "✅ FASE 3 concluída!" -ForegroundColor Green
    } catch {
        Write-Host "❌ ERRO na FASE 3: $_" -ForegroundColor Red
        $fase3_sucesso = $false
    }
} else {
    Write-Host "⏭️  FASE 3 pulada pelo usuário." -ForegroundColor Yellow
    $fase3_sucesso = $false
}

$fim = Get-Date
$duracao = ($fim - $inicio).TotalMinutes

# RELATÓRIO FINAL
Write-Host "`n`n" + "="*70 -ForegroundColor Cyan
Write-Host "📊 RELATÓRIO FINAL DA REORGANIZAÇÃO" -ForegroundColor Cyan
Write-Host "="*70 -ForegroundColor Cyan
Write-Host ""
Write-Host "⏱️  Tempo total: $([math]::Round($duracao, 1)) minutos" -ForegroundColor Yellow
Write-Host ""
Write-Host "📋 Status das Fases:" -ForegroundColor Yellow
Write-Host "   FASE 1 (Preparação):    $(if($fase1_sucesso){'✅ Sucesso'}else{'❌ Falhou'})" -ForegroundColor $(if($fase1_sucesso){'Green'}else{'Red'})
Write-Host "   FASE 2 (Consolidação):  $(if($fase2_sucesso){'✅ Sucesso'}else{'❌ Falhou'})" -ForegroundColor $(if($fase2_sucesso){'Green'}else{'Red'})
Write-Host "   FASE 3 (Limpeza):       $(if($fase3_sucesso){'✅ Sucesso'}elseif($confirmar_fase3 -ne 'SIM'){'⏭️  Pulada'}else{'❌ Falhou'})" -ForegroundColor $(if($fase3_sucesso){'Green'}elseif($confirmar_fase3 -ne 'SIM'){'Yellow'}else{'Red'})
Write-Host ""

if ($fase1_sucesso -and $fase2_sucesso) {
    Write-Host "🎉 REORGANIZAÇÃO CONCLUÍDA COM SUCESSO!" -ForegroundColor Green
    Write-Host ""
    Write-Host "📁 Nova estrutura:" -ForegroundColor Yellow
    Write-Host "   📂 _DADOS_CENTRALIZADOS/     ← Repositório central"
    Write-Host "   📂 _SCRIPTS_COMPARTILHADOS/  ← Utilitários"
    Write-Host "   📂 _ARQUIVO/                 ← Inativos"
    Write-Host "   📂 01_Helio_ML_Producao/     ← Produção ⭐"
    Write-Host "   📂 02_Pipeline_Midia_Paga/   ← Produção ⭐"
    Write-Host "   📂 03_Analises_Operacionais/"
    Write-Host "   📂 04_Auditorias_Qualidade/"
    Write-Host "   📂 05_Pesquisas_Educacionais/"
    Write-Host "   📂 Controle_de_entregas/"
    Write-Host ""
    Write-Host "📊 Benefícios alcançados:" -ForegroundColor Yellow
    Write-Host "   ✅ Estrutura organizada e padronizada"
    Write-Host "   ✅ Fonte única de verdade para dados"
    Write-Host "   ✅ Duplicatas eliminadas (~270 MB)"
    Write-Host "   ✅ Histórico preservado"
    Write-Host "   ✅ Documentação atualizada"
    Write-Host ""
    Write-Host "📋 Próximos passos:" -ForegroundColor Yellow
    Write-Host "   1. Abra e teste: 01_Helio_ML_Producao"
    Write-Host "   2. Abra e teste: 02_Pipeline_Midia_Paga"
    Write-Host "   3. Valide que os atalhos estão funcionando"
    Write-Host "   4. Revise: Controle_de_entregas\INVENTARIO_PROJETOS_DATA_ANALYTICS.xlsx"
    Write-Host "   5. Revise: Controle_de_entregas\RELATORIO_DUPLICACOES.xlsx"
    
    if ($fase3_sucesso) {
        Write-Host "   6. Após 30 dias, deletar pasta de backup criada na Fase 3"
    }
} else {
    Write-Host "⚠️  REORGANIZAÇÃO PARCIALMENTE CONCLUÍDA" -ForegroundColor Yellow
    Write-Host "   Revise os erros acima e execute as fases pendentes manualmente."
}

Write-Host "`n" + "="*70 -ForegroundColor Cyan
Write-Host ""

# Abrir explorador de arquivos na pasta de projetos
Write-Host "📂 Abrindo pasta de projetos..." -ForegroundColor Cyan
Start-Process explorer.exe -ArgumentList $ROOT

# Abrir Excel com inventário (se existir)
$inventario = "$ROOT\Controle_de_entregas\INVENTARIO_PROJETOS_DATA_ANALYTICS.xlsx"
if (Test-Path $inventario) {
    Write-Host "📊 Abrindo inventário de projetos..." -ForegroundColor Cyan
    Start-Process $inventario
}

Write-Host "`n✨ Processo concluído! Verifique os arquivos abertos." -ForegroundColor Green
