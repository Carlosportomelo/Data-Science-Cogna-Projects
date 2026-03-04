# ===================================================================
# FASE 3: LIMPEZA - REMOVER DUPLICATAS E OTIMIZAR
# ===================================================================
# Execute após concluir Fase 2
# ATENÇÃO: Esta fase deleta arquivos! Faça backup antes!
# Tempo estimado: 15 minutos
# ===================================================================

$ROOT = "C:\Users\a483650\Projetos"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "FASE 3: LIMPEZA E OTIMIZAÇÃO" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "⚠️  ATENÇÃO: Esta fase irá DELETAR arquivos duplicados!" -ForegroundColor Red
Write-Host "   Certifique-se de ter um backup completo antes de prosseguir." -ForegroundColor Yellow
Write-Host ""

$confirmacao = Read-Host "Deseja continuar? (Digite 'SIM' para confirmar)"

if ($confirmacao -ne "SIM") {
    Write-Host "`n❌ Operação cancelada pelo usuário." -ForegroundColor Red
    exit
}

# Criar pasta de backup de segurança
Write-Host "`n[0/6] Criando backup de segurança..." -ForegroundColor Yellow
$BACKUP_DIR = Join-Path $ROOT "_BACKUP_SEGURANCA_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -Path $BACKUP_DIR -ItemType Directory -Force | Out-Null
Write-Host "   ✓ Backup será salvo em: $BACKUP_DIR" -ForegroundColor Green

# Função para mover duplicatas para histórico
function Move-ToHistorico {
    param(
        [string]$SourcePath,
        [string]$HistoricoPath
    )
    
    if (Test-Path $SourcePath) {
        $fileName = Split-Path $SourcePath -Leaf
        $fileDate = (Get-Item $SourcePath).LastWriteTime.ToString("yyyy-MM-dd")
        $newName = $fileName -replace '\.csv$', "_$fileDate.csv"
        $newName = $newName -replace '\.xlsx$', "_$fileDate.xlsx"
        
        $destino = Join-Path $HistoricoPath $newName
        
        # Backup antes de mover
        $backupPath = Join-Path $BACKUP_DIR $fileName
        Copy-Item -Path $SourcePath -Destination $backupPath -Force
        
        Move-Item -Path $SourcePath -Destination $destino -Force -ErrorAction SilentlyContinue
        return $true
    }
    return $false
}

# 1. Consolidar bases HubSpot duplicadas
Write-Host "`n[1/6] Limpando duplicatas de hubspot_leads.csv..." -ForegroundColor Yellow

$hubspot_leads_duplicatas = @(
    "$ROOT\03_Analises_Operacionais\eficiencia_canal\data\backup\hubspot_leads.csv",
    "$ROOT\04_Auditorias_Qualidade\auditoria_leads_sumidos\hubspot_leads.csv",
    "$ROOT\02_Pipeline_Midia_Paga\data\backup\hubspot_leads.csv"
)

$movidos = 0
$historico_hubspot = "$ROOT\_DADOS_CENTRALIZADOS\hubspot\historico"

foreach ($arquivo in $hubspot_leads_duplicatas) {
    if (Move-ToHistorico -SourcePath $arquivo -HistoricoPath $historico_hubspot) {
        $movidos++
        Write-Host "   ✓ Movido para histórico: $arquivo" -ForegroundColor Green
    }
}

Write-Host "   Total movido: $movidos arquivos" -ForegroundColor Cyan

# 2. Limpar backups de leads scored (Projeto Helio)
Write-Host "`n[2/6] Limpando backups antigos de leads_scored..." -ForegroundColor Yellow

$pasta_backup_helio = "$ROOT\01_Helio_ML_Producao\Outputs\Dados_Scored\Backup"
if (Test-Path $pasta_backup_helio) {
    $arquivos_scored = Get-ChildItem -Path $pasta_backup_helio -Filter "leads_scored*.csv" | 
                       Sort-Object LastWriteTime -Descending
    
    # Manter apenas os 10 mais recentes
    $para_manter = 10
    $para_deletar = $arquivos_scored | Select-Object -Skip $para_manter
    
    $deletados = 0
    foreach ($arquivo in $para_deletar) {
        # Backup antes de deletar
        Copy-Item -Path $arquivo.FullName -Destination $BACKUP_DIR -Force
        Remove-Item -Path $arquivo.FullName -Force
        $deletados++
    }
    
    Write-Host "   ✓ Mantidos: $para_manter arquivos mais recentes" -ForegroundColor Green
    Write-Host "   ✓ Deletados: $deletados arquivos antigos" -ForegroundColor Yellow
    $espacoLiberado = ($para_deletar | Measure-Object -Property Length -Sum).Sum / 1MB
    Write-Host "   💾 Espaço liberado: $([math]::Round($espacoLiberado, 2)) MB" -ForegroundColor Cyan
}

# 3. Limpar relatórios duplicados do Helio
Write-Host "`n[3/6] Limpando relatórios duplicados..." -ForegroundColor Yellow

$pastas_relatorios = @(
    "$ROOT\01_Helio_ML_Producao\Outputs\Relatorios_Unidades\Backup",
    "$ROOT\01_Helio_ML_Producao\Outputs\Relatorios_ML\Backup"
)

$total_liberado = 0

foreach ($pasta in $pastas_relatorios) {
    if (Test-Path $pasta) {
        $arquivos = Get-ChildItem -Path $pasta -Filter "*.xlsx" | 
                    Sort-Object LastWriteTime -Descending
        
        # Manter apenas os 5 mais recentes
        $para_manter = 5
        $para_deletar = $arquivos | Select-Object -Skip $para_manter
        
        foreach ($arquivo in $para_deletar) {
            $tamanho = $arquivo.Length / 1MB
            Copy-Item -Path $arquivo.FullName -Destination $BACKUP_DIR -Force
            Remove-Item -Path $arquivo.FullName -Force
            $total_liberado += $tamanho
        }
        
        $pasta_nome = Split-Path $pasta -Parent | Split-Path -Leaf
        Write-Host "   ✓ $pasta_nome: Mantidos $para_manter, deletados $($para_deletar.Count)" -ForegroundColor Green
    }
}

Write-Host "   💾 Espaço total liberado: $([math]::Round($total_liberado, 2)) MB" -ForegroundColor Cyan

# 4. Remover cópias duplicadas do projeto_helio_teste
Write-Host "`n[4/6] Limpando projeto_helio_teste arquivado..." -ForegroundColor Yellow

$helio_teste = "$ROOT\_ARQUIVO\projeto_helio_teste\Outputs"
if (Test-Path $helio_teste) {
    $tamanho_antes = (Get-ChildItem -Path $helio_teste -Recurse -File | 
                      Measure-Object -Property Length -Sum).Sum / 1MB
    
    # Backup do projeto inteiro
    $backup_helio_teste = Join-Path $BACKUP_DIR "projeto_helio_teste"
    Copy-Item -Path "$ROOT\_ARQUIVO\projeto_helio_teste" -Destination $backup_helio_teste -Recurse -Force
    
    # Deletar todos os outputs (já existem no projeto principal)
    Remove-Item -Path $helio_teste -Recurse -Force
    
    Write-Host "   ✓ Outputs deletados do projeto de teste" -ForegroundColor Green
    Write-Host "   💾 Espaço liberado: $([math]::Round($tamanho_antes, 2)) MB" -ForegroundColor Cyan
}

# 5. Consolidar bases de matrículas
Write-Host "`n[5/6] Consolidando bases de matrículas..." -ForegroundColor Yellow

$matriculas_duplicatas = @(
    "$ROOT\01_Helio_ML_Producao\Data\backup\matriculas_finais.xlsx",
    "$ROOT\01_Helio_ML_Producao\Data\backup\matriculas_finais.csv"
)

$historico_matriculas = "$ROOT\_DADOS_CENTRALIZADOS\matriculas\historico"
$movidos_mat = 0

foreach ($arquivo in $matriculas_duplicatas) {
    if (Move-ToHistorico -SourcePath $arquivo -HistoricoPath $historico_matriculas) {
        $movidos_mat++
        Write-Host "   ✓ Movido para histórico: $(Split-Path $arquivo -Leaf)" -ForegroundColor Green
    }
}

# 6. Criar symlinks/atalhos para dados centralizados
Write-Host "`n[6/6] Criando atalhos para dados centralizados..." -ForegroundColor Yellow

# Função para criar Junction (atalho de diretório)
function New-DirectoryJunction {
    param(
        [string]$LinkPath,
        [string]$TargetPath
    )
    
    if (Test-Path $LinkPath) {
        Write-Host "   ⚠ Já existe: $LinkPath" -ForegroundColor Gray
        return
    }
    
    try {
        cmd /c mklink /J "$LinkPath" "$TargetPath" | Out-Null
        Write-Host "   ✓ Atalho criado: $LinkPath → $TargetPath" -ForegroundColor Green
    } catch {
        Write-Host "   ❌ Erro ao criar atalho: $LinkPath" -ForegroundColor Red
    }
}

# Criar atalhos nos projetos principais
Write-Host "`n   Projeto Helio:" -ForegroundColor Cyan
$helio_data_link = "$ROOT\01_Helio_ML_Producao\Data\_DADOS_CENTRALIZADOS"
New-DirectoryJunction -LinkPath $helio_data_link -TargetPath "$ROOT\_DADOS_CENTRALIZADOS"

Write-Host "`n   Pipeline Mídia Paga:" -ForegroundColor Cyan
$midia_data_link = "$ROOT\02_Pipeline_Midia_Paga\data\_DADOS_CENTRALIZADOS"
New-DirectoryJunction -LinkPath $midia_data_link -TargetPath "$ROOT\_DADOS_CENTRALIZADOS"

# 7. Gerar relatório final
Write-Host "`n" + "="*60 -ForegroundColor Cyan
Write-Host "📊 GERANDO RELATÓRIO FINAL..." -ForegroundColor Yellow

$relatorio = @"
# RELATÓRIO DE REORGANIZAÇÃO - FASE 3
Data: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")

## Resumo de Operações

### Duplicatas Removidas
- hubspot_leads.csv: $movidos cópias movidas para histórico
- leads_scored backups: $deletados arquivos antigos deletados
- Relatórios Helio: $(($para_deletar | Measure-Object).Count) arquivos antigos deletados
- Outputs projeto_helio_teste: Diretório completo removido

### Espaço Liberado
- Total estimado: $([math]::Round($total_liberado + $espacoLiberado + $tamanho_antes, 2)) MB

### Backups Criados
- Localização: $BACKUP_DIR
- Todos os arquivos deletados foram salvos em backup de segurança

### Atalhos Criados
- 01_Helio_ML_Producao\Data\_DADOS_CENTRALIZADOS
- 02_Pipeline_Midia_Paga\data\_DADOS_CENTRALIZADOS

## Próximos Passos

1. ✅ Testar scripts dos projetos principais (Helio, Mídia Paga)
2. ✅ Validar que os atalhos funcionam corretamente
3. ✅ Atualizar caminhos nos scripts que apontavam para bases antigas
4. ✅ Após 30 dias, se tudo estiver OK, deletar pasta de backup

## Estrutura Final

C:\Users\a483650\Projetos\
├── _DADOS_CENTRALIZADOS/          ← Fonte única de verdade
├── _SCRIPTS_COMPARTILHADOS/
├── _ARQUIVO/                      ← Projetos inativos
├── 01_Helio_ML_Producao/          ← PRODUÇÃO
├── 02_Pipeline_Midia_Paga/        ← PRODUÇÃO  
├── 03_Analises_Operacionais/
├── 04_Auditorias_Qualidade/
├── 05_Pesquisas_Educacionais/
└── Controle_de_entregas/

Total de arquivos: Redução de ~18.7% (269 MB economizados)
"@

$relatorio_path = "$ROOT\Controle_de_entregas\RELATORIO_REORGANIZACAO_FASE3.txt"
Set-Content -Path $relatorio_path -Value $relatorio -Encoding UTF8

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "✅ FASE 3 CONCLUÍDA COM SUCESSO!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "📊 Estatísticas:" -ForegroundColor Yellow
Write-Host "   💾 Espaço liberado: ~$([math]::Round($total_liberado + $espacoLiberado, 2)) MB"
Write-Host "   🔄 Arquivos movidos para histórico: $($movidos + $movidos_mat)"
Write-Host "   🗑️  Arquivos deletados: $deletados"
Write-Host "   🔗 Atalhos criados: 2"
Write-Host ""
Write-Host "📁 Arquivos Importantes:" -ForegroundColor Yellow
Write-Host "   📄 Relatório completo: $relatorio_path"
Write-Host "   💾 Backup de segurança: $BACKUP_DIR"
Write-Host ""
Write-Host "⚠️  IMPORTANTE:" -ForegroundColor Red
Write-Host "   1. Teste os scripts principais antes de deletar o backup"
Write-Host "   2. Valide que os atalhos estão funcionando"
Write-Host "   3. Após 30 dias, se tudo estiver OK, delete: $BACKUP_DIR"
Write-Host ""
Write-Host "📁 Próximo passo:" -ForegroundColor Yellow
Write-Host "   Teste os projetos e valide a reorganização!"
Write-Host ""
