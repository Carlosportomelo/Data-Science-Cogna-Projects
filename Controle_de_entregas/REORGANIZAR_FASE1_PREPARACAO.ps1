# ===================================================================
# FASE 1: PREPARAÇÃO - CRIAR ESTRUTURA CENTRALIZADA
# ===================================================================
# Execute este script primeiro para criar a estrutura base
# Tempo estimado: 5 minutos
# ===================================================================

$ROOT = "C:\Users\a483650\Projetos"
$TIMESTAMP = Get-Date -Format "yyyy-MM-dd_HHmmss"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "FASE 1: PREPARAÇÃO" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 1. Criar estrutura _DADOS_CENTRALIZADOS
Write-Host "[1/4] Criando estrutura _DADOS_CENTRALIZADOS..." -ForegroundColor Yellow

$estrutura = @(
    "_DADOS_CENTRALIZADOS",
    "_DADOS_CENTRALIZADOS\hubspot",
    "_DADOS_CENTRALIZADOS\hubspot\historico",
    "_DADOS_CENTRALIZADOS\matriculas",
    "_DADOS_CENTRALIZADOS\matriculas\historico",
    "_DADOS_CENTRALIZADOS\marketing",
    "_DADOS_CENTRALIZADOS\marketing\historico"
)

foreach ($pasta in $estrutura) {
    $caminho = Join-Path $ROOT $pasta
    if (!(Test-Path $caminho)) {
        New-Item -Path $caminho -ItemType Directory -Force | Out-Null
        Write-Host "   ✓ Criado: $pasta" -ForegroundColor Green
    } else {
        Write-Host "   ⚠ Já existe: $pasta" -ForegroundColor Gray
    }
}

# 2. Criar estrutura _SCRIPTS_COMPARTILHADOS
Write-Host "`n[2/4] Criando estrutura _SCRIPTS_COMPARTILHADOS..." -ForegroundColor Yellow

$estrutura_scripts = @(
    "_SCRIPTS_COMPARTILHADOS",
    "_SCRIPTS_COMPARTILHADOS\utils"
)

foreach ($pasta in $estrutura_scripts) {
    $caminho = Join-Path $ROOT $pasta
    if (!(Test-Path $caminho)) {
        New-Item -Path $caminho -ItemType Directory -Force | Out-Null
        Write-Host "   ✓ Criado: $pasta" -ForegroundColor Green
    }
}

# 3. Criar pasta _ARQUIVO para projetos inativos
Write-Host "`n[3/4] Criando pasta _ARQUIVO..." -ForegroundColor Yellow

$pasta_arquivo = Join-Path $ROOT "_ARQUIVO"
if (!(Test-Path $pasta_arquivo)) {
    New-Item -Path $pasta_arquivo -ItemType Directory -Force | Out-Null
    Write-Host "   ✓ Criado: _ARQUIVO" -ForegroundColor Green
}

# 4. Copiar bases mais recentes para repositório central
Write-Host "`n[4/4] Copiando bases mais recentes..." -ForegroundColor Yellow

# Identificar e copiar hubspot_leads.csv mais recente
Write-Host "   📄 Processando hubspot_leads.csv..." -ForegroundColor Cyan
$hubspot_leads_files = @(
    "$ROOT\analise_eficiencia_canal\data\hubspot_leads.csv",
    "$ROOT\projeto_helio\Data\hubspot_leads.csv",
    "$ROOT\analise_performance_midiapaga\data\hubspot_leads.csv"
)

$mais_recente_leads = $null
$data_mais_recente = [DateTime]::MinValue

foreach ($file in $hubspot_leads_files) {
    if (Test-Path $file) {
        $ultima_modificacao = (Get-Item $file).LastWriteTime
        if ($ultima_modificacao -gt $data_mais_recente) {
            $data_mais_recente = $ultima_modificacao
            $mais_recente_leads = $file
        }
    }
}

if ($mais_recente_leads) {
    $destino = "$ROOT\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_ATUAL.csv"
    Copy-Item -Path $mais_recente_leads -Destination $destino -Force
    $tamanho = [math]::Round((Get-Item $destino).Length / 1MB, 2)
    Write-Host "      ✓ Copiado: hubspot_leads_ATUAL.csv ($tamanho MB)" -ForegroundColor Green
    Write-Host "      Fonte: $mais_recente_leads" -ForegroundColor Gray
    Write-Host "      Data: $data_mais_recente" -ForegroundColor Gray
}

# Copiar hubspot_negocios_perdidos.csv
Write-Host "`n   📄 Processando hubspot_negocios_perdidos.csv..." -ForegroundColor Cyan
$hubspot_perdidos = "$ROOT\projeto_helio\Data\hubspot_negocios_perdidos.csv"
if (Test-Path $hubspot_perdidos) {
    $destino = "$ROOT\_DADOS_CENTRALIZADOS\hubspot\hubspot_negocios_perdidos_ATUAL.csv"
    Copy-Item -Path $hubspot_perdidos -Destination $destino -Force
    $tamanho = [math]::Round((Get-Item $destino).Length / 1MB, 2)
    Write-Host "      ✓ Copiado: hubspot_negocios_perdidos_ATUAL.csv ($tamanho MB)" -ForegroundColor Green
}

# Copiar matriculas_finais
Write-Host "`n   📄 Processando matriculas_finais..." -ForegroundColor Cyan
$matriculas_xlsx = "$ROOT\projeto_helio\Data\matriculas_finais.xlsx"
$matriculas_csv = "$ROOT\projeto_helio\Data\matriculas_finais_limpo.csv"

if (Test-Path $matriculas_xlsx) {
    Copy-Item -Path $matriculas_xlsx -Destination "$ROOT\_DADOS_CENTRALIZADOS\matriculas\matriculas_finais_ATUAL.xlsx" -Force
    Write-Host "      ✓ Copiado: matriculas_finais_ATUAL.xlsx" -ForegroundColor Green
}
if (Test-Path $matriculas_csv) {
    Copy-Item -Path $matriculas_csv -Destination "$ROOT\_DADOS_CENTRALIZADOS\matriculas\matriculas_finais_ATUAL.csv" -Force
    Write-Host "      ✓ Copiado: matriculas_finais_ATUAL.csv" -ForegroundColor Green
}

# Copiar bases de marketing
Write-Host "`n   📄 Processando bases de marketing..." -ForegroundColor Cyan
$meta_dataset = "$ROOT\analise_performance_midiapaga\data\meta_dataset.csv"
if (Test-Path $meta_dataset) {
    Copy-Item -Path $meta_dataset -Destination "$ROOT\_DADOS_CENTRALIZADOS\marketing\meta_ads_ATUAL.csv" -Force
    Write-Host "      ✓ Copiado: meta_ads_ATUAL.csv" -ForegroundColor Green
}

# Criar README.md de documentação
Write-Host "`n   📝 Criando documentação..." -ForegroundColor Cyan
$readme_content = @"
# Repositório Central de Dados

Este diretório contém as versões mais recentes e oficiais de todas as bases de dados utilizadas nos projetos.

## 📋 Estrutura

### hubspot/
- **hubspot_leads_ATUAL.csv**: Base principal de leads do HubSpot CRM
- **hubspot_negocios_perdidos_ATUAL.csv**: Negócios que não converteram
- **historico/**: Versões anteriores organizadas por data

### matriculas/
- **matriculas_finais_ATUAL.csv**: Base limpa de matrículas confirmadas
- **matriculas_finais_ATUAL.xlsx**: Versão Excel com formatação
- **historico/**: Versões anteriores

### marketing/
- **meta_ads_ATUAL.csv**: Dados de campanhas Meta Ads (Facebook/Instagram)
- **google_ads_ATUAL.csv**: Dados de campanhas Google Ads
- **historico/**: Histórico de campanhas

## 🔄 Atualização

Quando atualizar uma base:
1. Copie a versão atual para /historico/ com sufixo de data
2. Substitua o arquivo _ATUAL.csv pela nova versão
3. Documente a atualização abaixo

## 📊 Última Atualização

- **Data**: $(Get-Date -Format "dd/MM/yyyy HH:mm")
- **Responsável**: Script automático de reorganização
- **Ação**: Criação inicial do repositório central

## 🔗 Projetos que Usam Estas Bases

- **01_Helio_ML_Producao**: Todas as bases
- **02_Pipeline_Midia_Paga**: hubspot_leads, meta_ads
- **03_Analises_Operacionais**: Conforme necessidade de cada análise
- **04_Auditorias_Qualidade**: hubspot_leads, matriculas

"@

Set-Content -Path "$ROOT\_DADOS_CENTRALIZADOS\README.md" -Value $readme_content -Encoding UTF8
Write-Host "      ✓ Criado: README.md" -ForegroundColor Green

# Resumo final
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "✅ FASE 1 CONCLUÍDA COM SUCESSO!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "📊 Resumo:" -ForegroundColor Yellow
Write-Host "   ✓ Estrutura _DADOS_CENTRALIZADOS criada"
Write-Host "   ✓ Estrutura _SCRIPTS_COMPARTILHADOS criada"
Write-Host "   ✓ Pasta _ARQUIVO criada"
Write-Host "   ✓ Bases principais copiadas"
Write-Host "   ✓ Documentação gerada"
Write-Host ""
Write-Host "📁 Próximo passo:" -ForegroundColor Yellow
Write-Host "   Execute: .\REORGANIZAR_FASE2_CONSOLIDACAO.ps1"
Write-Host ""
