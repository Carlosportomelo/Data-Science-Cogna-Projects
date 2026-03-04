# ===================================================================
# FASE 2: CONSOLIDAÇÃO - REORGANIZAR PROJETOS  
# ===================================================================
# Execute após concluir Fase 1
# Tempo estimado: 10 minutos
# ===================================================================

$ROOT = "C:\Users\a483650\Projetos"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "FASE 2: CONSOLIDAÇÃO" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 1. Renomear projetos principais (Produção)
Write-Host "[1/5] Renomeando projetos de produção..." -ForegroundColor Yellow

$renomeacoes_producao = @{
    "projeto_helio" = "01_Helio_ML_Producao"
    "analise_performance_midiapaga" = "02_Pipeline_Midia_Paga"
}

foreach ($antigo in $renomeacoes_producao.Keys) {
    $novo = $renomeacoes_producao[$antigo]
    $caminho_antigo = Join-Path $ROOT $antigo
    $caminho_novo = Join-Path $ROOT $novo
    
    if (Test-Path $caminho_antigo) {
        if (!(Test-Path $caminho_novo)) {
            Rename-Item -Path $caminho_antigo -NewName $novo
            Write-Host "   ✓ Renomeado: $antigo → $novo" -ForegroundColor Green
        } else {
            Write-Host "   ⚠ Já existe: $novo (pulando)" -ForegroundColor Yellow
        }
    } else {
        Write-Host "   ⚠ Não encontrado: $antigo" -ForegroundColor Gray
    }
}

# 2. Criar estrutura consolidada de Análises Operacionais
Write-Host "`n[2/5] Criando estrutura 03_Analises_Operacionais..." -ForegroundColor Yellow

$pasta_analises = Join-Path $ROOT "03_Analises_Operacionais"
if (!(Test-Path $pasta_analises)) {
    New-Item -Path $pasta_analises -ItemType Directory -Force | Out-Null
    Write-Host "   ✓ Criado: 03_Analises_Operacionais" -ForegroundColor Green
}

$projetos_operacionais = @{
    "analise_eficiencia_canal" = "eficiencia_canal"
    "Comparativo unidades" = "comparativo_unidades"
    "análise_curva_alunos" = "curva_alunos"
    "Analise_performance_valor_the_news" = "valor_the_news"
}

foreach ($antigo in $projetos_operacionais.Keys) {
    $novo = $projetos_operacionais[$antigo]
    $caminho_antigo = Join-Path $ROOT $antigo
    $caminho_novo = Join-Path $pasta_analises $novo
    
    if (Test-Path $caminho_antigo) {
        if (!(Test-Path $caminho_novo)) {
            Move-Item -Path $caminho_antigo -Destination $caminho_novo -Force
            Write-Host "   ✓ Movido: $antigo → 03_Analises_Operacionais\$novo" -ForegroundColor Green
        } else {
            Write-Host "   ⚠ Já existe em destino: $novo" -ForegroundColor Yellow
        }
    }
}

# 3. Criar estrutura consolidada de Auditorias
Write-Host "`n[3/5] Criando estrutura 04_Auditorias_Qualidade..." -ForegroundColor Yellow

$pasta_auditorias = Join-Path $ROOT "04_Auditorias_Qualidade"
if (!(Test-Path $pasta_auditorias)) {
    New-Item -Path $pasta_auditorias -ItemType Directory -Force | Out-Null
    Write-Host "   ✓ Criado: 04_Auditorias_Qualidade" -ForegroundColor Green
}

$projetos_auditoria = @{
    "Audiotira_de_bases" = "auditoria_leads_sumidos"
    "correção_preenchimento_meta_callcenter" = "correcao_meta_callcenter"
}

foreach ($antigo in $projetos_auditoria.Keys) {
    $novo = $projetos_auditoria[$antigo]
    $caminho_antigo = Join-Path $ROOT $antigo
    $caminho_novo = Join-Path $pasta_auditorias $novo
    
    if (Test-Path $caminho_antigo) {
        if (!(Test-Path $caminho_novo)) {
            Move-Item -Path $caminho_antigo -Destination $caminho_novo -Force
            Write-Host "   ✓ Movido: $antigo → 04_Auditorias_Qualidade\$novo" -ForegroundColor Green
        }
    }
}

# 4. Criar estrutura de Pesquisas Educacionais
Write-Host "`n[4/5] Criando estrutura 05_Pesquisas_Educacionais..." -ForegroundColor Yellow

$pasta_pesquisas = Join-Path $ROOT "05_Pesquisas_Educacionais"
if (!(Test-Path $pasta_pesquisas)) {
    New-Item -Path $pasta_pesquisas -ItemType Directory -Force | Out-Null
    Write-Host "   ✓ Criado: 05_Pesquisas_Educacionais" -ForegroundColor Green
}

$projetos_pesquisa = @(
    "Analise_Cultura_Inglesa_CEFR",
    "Pesquisa_Correlacao_CEFR_Cultura_inglesa"
)

foreach ($projeto in $projetos_pesquisa) {
    $caminho_antigo = Join-Path $ROOT $projeto
    $caminho_novo = Join-Path $pasta_pesquisas $projeto
    
    if (Test-Path $caminho_antigo) {
        if (!(Test-Path $caminho_novo)) {
            Move-Item -Path $caminho_antigo -Destination $caminho_novo -Force
            Write-Host "   ✓ Movido: $projeto → 05_Pesquisas_Educacionais\" -ForegroundColor Green
        }
    }
}

# 5. Mover projetos inativos para _ARQUIVO
Write-Host "`n[5/5] Movendo projetos inativos para _ARQUIVO..." -ForegroundColor Yellow

$projetos_arquivo = @(
    "AMBIENTE_TESTE_ISOLADO_2025-12-15",
    "projeto_helio_teste",
    "analise_performance_rvo",
    "analise_performance_midia_organica",
    "apresentação_radar"
)

$pasta_arquivo = Join-Path $ROOT "_ARQUIVO"

foreach ($projeto in $projetos_arquivo) {
    $caminho_antigo = Join-Path $ROOT $projeto
    $caminho_novo = Join-Path $pasta_arquivo $projeto
    
    if (Test-Path $caminho_antigo) {
        if (!(Test-Path $caminho_novo)) {
            Move-Item -Path $caminho_antigo -Destination $caminho_novo -Force
            Write-Host "   ✓ Arquivado: $projeto" -ForegroundColor Green
        }
    }
}

# Criar arquivo de índice na pasta _ARQUIVO
$index_content = @"
# Projetos Arquivados

Este diretório contém projetos inativos, descontinuados ou de teste.

## Lista de Projetos

### AMBIENTE_TESTE_ISOLADO_2025-12-15
- **Status**: Ambiente de teste
- **Descrição**: Sandbox para testes isolados
- **Data de arquivamento**: $(Get-Date -Format "dd/MM/yyyy")

### projeto_helio_teste  
- **Status**: Versão de desenvolvimento
- **Descrição**: Ambiente de teste do Projeto Helio (PRJ-001)
- **Data de arquivamento**: $(Get-Date -Format "dd/MM/yyyy")

### analise_performance_rvo
- **Status**: Projeto vazio/planejado
- **Descrição**: Pasta vazia
- **Data de arquivamento**: $(Get-Date -Format "dd/MM/yyyy")

### analise_performance_midia_organica
- **Status**: Projeto vazio/planejado
- **Descrição**: Pasta vazia
- **Data de arquivamento**: $(Get-Date -Format "dd/MM/yyyy")

### apresentação_radar
- **Status**: Concluído
- **Descrição**: Templates HTML para visualizações Instagram
- **Data de arquivamento**: $(Get-Date -Format "dd/MM/yyyy")

## Reativação

Para reativar um projeto, mova-o de volta para a pasta raiz de Projetos.
"@

Set-Content -Path "$pasta_arquivo\README.md" -Value $index_content -Encoding UTF8
Write-Host "   ✓ Criado: _ARQUIVO\README.md" -ForegroundColor Green

# Resumo final
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "✅ FASE 2 CONCLUÍDA COM SUCESSO!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "📊 Resumo:" -ForegroundColor Yellow
Write-Host "   ✓ Projetos de produção renomeados"
Write-Host "   ✓ Análises operacionais consolidadas"
Write-Host "   ✓ Auditorias consolidadas"
Write-Host "   ✓ Pesquisas organizadas"
Write-Host "   ✓ Projetos inativos arquivados"
Write-Host ""
Write-Host "📁 Estrutura atual:" -ForegroundColor Yellow
Write-Host "   📂 01_Helio_ML_Producao (Produção)"
Write-Host "   📂 02_Pipeline_Midia_Paga (Produção)"
Write-Host "   📂 03_Analises_Operacionais (4 projetos)"
Write-Host "   📂 04_Auditorias_Qualidade (2 projetos)"
Write-Host "   📂 05_Pesquisas_Educacionais"
Write-Host "   📂 _DADOS_CENTRALIZADOS"
Write-Host "   📂 _SCRIPTS_COMPARTILHADOS"
Write-Host "   📂 _ARQUIVO (5 projetos)"
Write-Host ""
Write-Host "📁 Próximo passo:" -ForegroundColor Yellow
Write-Host "   Execute: .\REORGANIZAR_FASE3_LIMPEZA.ps1"
Write-Host ""
