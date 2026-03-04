# 🏗️ GUIA DE REORGANIZAÇÃO DE PROJETOS

**Data de Criação**: 09/02/2026  
**Projeto**: Reorganização do Diretório C:\Users\a483650\Projetos  
**Objetivo**: Eliminar duplicações, padronizar estrutura e centralizar dados

---

## 📋 RESUMO EXECUTIVO

### Situação Atual (Antes da Reorganização)
- **Total de arquivos**: 1.581 (CSV + XLSX)
- **Espaço ocupado**: 1.42 GB
- **Duplicações identificadas**: 212 grupos de arquivos idênticos
- **Espaço desperdiçado**: ~270 MB (18.7%)
- **Problemas principais**:
  - 6 cópias diferentes de `hubspot_leads.csv`
  - Backups espalhados sem organização
  - Projetos sem nomenclatura padronizada
  - Bases críticas duplicadas em múltiplos projetos

### Estrutura Proposta (Após Reorganização)
```
C:\Users\a483650\Projetos\
│
├── 📂 _DADOS_CENTRALIZADOS/          ⭐ NOVO - Fonte única de verdade
│   ├── hubspot/
│   │   ├── hubspot_leads_ATUAL.csv
│   │   ├── hubspot_negocios_perdidos_ATUAL.csv
│   │   └── historico/
│   ├── matriculas/
│   │   ├── matriculas_finais_ATUAL.csv
│   │   └── historico/
│   ├── marketing/
│   │   ├── meta_ads_ATUAL.csv
│   │   └── historico/
│   └── README.md
│
├── 📂 _SCRIPTS_COMPARTILHADOS/       ⭐ NOVO - Utilitários reutilizáveis
│   └── utils/
│
├── 📂 01_Helio_ML_Producao/          ✅ PRODUÇÃO - Renomeado
├── 📂 02_Pipeline_Midia_Paga/        ✅ PRODUÇÃO - Renomeado
├── 📂 03_Analises_Operacionais/      ⭐ CONSOLIDADO
│   ├── eficiencia_canal/
│   ├── comparativo_unidades/
│   ├── curva_alunos/
│   └── valor_the_news/
│
├── 📂 04_Auditorias_Qualidade/       ⭐ CONSOLIDADO
│   ├── auditoria_leads_sumidos/
│   └── correcao_meta_callcenter/
│
├── 📂 05_Pesquisas_Educacionais/     ⭐ CONSOLIDADO
│
├── 📂 _ARQUIVO/                      ⭐ NOVO - Projetos inativos
│   ├── ambiente_teste/
│   ├── projeto_helio_teste/
│   └── projetos_vazios/
│
└── 📂 Controle_de_entregas/          ✅ MANTIDO
    ├── INVENTARIO_PROJETOS_DATA_ANALYTICS.xlsx
    ├── RELATORIO_DUPLICACOES.xlsx
    └── Scripts de reorganização
```

---

## 📦 ARQUIVOS GERADOS

### 1. Análises e Inventários
| Arquivo | Descrição |
|---------|-----------|
| `INVENTARIO_PROJETOS_DATA_ANALYTICS.xlsx` | Catálogo completo de 15 projetos com metadados |
| `RELATORIO_DUPLICACOES.xlsx` | Top 20 arquivos duplicados e análise de bases críticas |
| `analisar_duplicacoes.py` | Script Python que gerou as análises |
| `inventario_projetos.py` | Script Python que gerou o inventário |

### 2. Scripts de Reorganização (PowerShell)
| Arquivo | Fase | Descrição |
|---------|------|-----------|
| `REORGANIZAR_MASTER.ps1` | **Master** | Executa todas as fases automaticamente |
| `REORGANIZAR_FASE1_PREPARACAO.ps1` | Fase 1 | Cria estrutura e copia bases centralizadas |
| `REORGANIZAR_FASE2_CONSOLIDACAO.ps1` | Fase 2 | Reorganiza e renomeia projetos |
| `REORGANIZAR_FASE3_LIMPEZA.ps1` | Fase 3 | Remove duplicatas e otimiza espaço |

---

## 🚀 COMO EXECUTAR A REORGANIZAÇÃO

### ⚠️ ANTES DE COMEÇAR

1. **BACKUP OBRIGATÓRIO**: Faça backup completo da pasta `C:\Users\a483650\Projetos`
2. **Feche todos os programas**: Excel, Python, VS Code, etc.
3. **Permissões**: Execute PowerShell como Administrador
4. **Tempo necessário**: 30-45 minutos

### Opção 1: Execução Automática (Recomendado)

```powershell
# Abrir PowerShell como Administrador
cd "C:\Users\a483650\Projetos\Controle_de_entregas"

# Permitir execução de scripts (se necessário)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Executar MASTER (todas as fases)
.\REORGANIZAR_MASTER.ps1
```

### Opção 2: Execução Manual por Fases

Se preferir executar passo a passo com mais controle:

```powershell
# FASE 1: Preparação (5 min) - Sem riscos
.\REORGANIZAR_FASE1_PREPARACAO.ps1

# Pausa para revisar
# Verifique se a pasta _DADOS_CENTRALIZADOS foi criada corretamente

# FASE 2: Consolidação (10 min) - Movimenta pastas
.\REORGANIZAR_FASE2_CONSOLIDACAO.ps1

# Pausa para revisar
# Verifique se os projetos foram renomeados/movidos corretamente

# FASE 3: Limpeza (15 min) - DELETA ARQUIVOS!
# ⚠️ ATENÇÃO: Esta fase remove duplicatas permanentemente
.\REORGANIZAR_FASE3_LIMPEZA.ps1
```

---

## 📊 O QUE CADA FASE FAZ

### FASE 1: Preparação (Segura)
✅ Cria estrutura `_DADOS_CENTRALIZADOS/`  
✅ Cria estrutura `_SCRIPTS_COMPARTILHADOS/`  
✅ Cria pasta `_ARQUIVO/`  
✅ **Copia** (não move) bases mais recentes:
   - `hubspot_leads.csv` (mais recente de 3 locais)
   - `hubspot_negocios_perdidos.csv`
   - `matriculas_finais.csv`
   - `meta_ads.csv`
✅ Gera documentação README.md

**Riscos**: ❌ Nenhum (só cria e copia)

---

### FASE 2: Consolidação (Movimenta)
✅ Renomeia projetos principais:
   - `projeto_helio` → `01_Helio_ML_Producao`
   - `analise_performance_midiapaga` → `02_Pipeline_Midia_Paga`
✅ Consolida análises em `03_Analises_Operacionais/`
✅ Consolida auditorias em `04_Auditorias_Qualidade/`
✅ Agrupa pesquisas em `05_Pesquisas_Educacionais/`
✅ Move projetos inativos para `_ARQUIVO/`

**Riscos**: ⚠️ Baixo (move pastas, mas não deleta)

---

### FASE 3: Limpeza (DELETA arquivos!)
⚠️ **Move duplicatas de hubspot_leads.csv para histórico**  
⚠️ **Deleta backups antigos** (mantém 10 mais recentes)  
⚠️ **Remove outputs duplicados** do projeto_helio_teste  
✅ **Cria backup de segurança completo antes de deletar**  
✅ Cria atalhos (junctions) para `_DADOS_CENTRALIZADOS/`  
✅ Gera relatório detalhado

**Riscos**: ⚠️⚠️ ALTO (deleta arquivos, mas cria backup antes)

**Espaço liberado**: ~270 MB

---

## 🎯 BENEFÍCIOS DA REORGANIZAÇÃO

### 1. Eliminação de Duplicações
- ❌ **Antes**: 6 cópias de `hubspot_leads.csv` (29 MB cada = 174 MB)
- ✅ **Depois**: 1 versão oficial + histórico organizado

### 2. Fonte Única de Verdade
- Todos os projetos apontam para `_DADOS_CENTRALIZADOS/`
- Atualiza uma vez, todos os projetos usam a mesma base

### 3. Estrutura Padronizada
- Nomenclatura clara: `01_`, `02_`, `03_`...
- Categoria óbvia: Produção, Análises, Auditorias, etc.

### 4. Histórico Preservado
- Backups organizados por data em `/historico/`
- Rastreabilidade de mudanças

### 5. Redução de Espaço
- **Antes**: 1.42 GB
- **Depois**: ~1.15 GB
- **Economia**: 270 MB (18.7%)

### 6. Manutenibilidade
- Fácil localizar projetos
- Scripts compartilhados reutilizáveis
- Documentação centralizada

---

## 🔍 VALIDAÇÃO PÓS-REORGANIZAÇÃO

### Checklist de Testes

```powershell
# 1. Verificar estrutura criada
cd "C:\Users\a483650\Projetos"
dir

# Deve conter:
# _DADOS_CENTRALIZADOS, _SCRIPTS_COMPARTILHADOS, _ARQUIVO
# 01_Helio_ML_Producao, 02_Pipeline_Midia_Paga
# 03_Analises_Operacionais, 04_Auditorias_Qualidade
# 05_Pesquisas_Educacionais, Controle_de_entregas
```

### Testar Projeto Helio

```powershell
cd "C:\Users\a483650\Projetos\01_Helio_ML_Producao"

# Verificar que pasta Data existe
dir Data

# Verificar atalho para dados centralizados
dir Data\_DADOS_CENTRALIZADOS

# Testar script principal (se configurado)
# python Scritps\1.ML_Lead_Scoring.py
```

### Testar Pipeline Mídia Paga

```powershell
cd "C:\Users\a483650\Projetos\02_Pipeline_Midia_Paga"

# Verificar atalho
dir data\_DADOS_CENTRALIZADOS

# Bases devem estar acessíveis via atalho
```

### Validar Dados Centralizados

```powershell
cd "C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS"

# Verificar bases principais
dir hubspot\hubspot_leads_ATUAL.csv
dir matriculas\matriculas_finais_ATUAL.csv
dir marketing\meta_ads_ATUAL.csv

# Verificar históricos
dir hubspot\historico
dir matriculas\historico
```

---

## 📋 ATUALIZAÇÃO DE BASES (Pós-Reorganização)

### Processo Padrão para Atualizar uma Base

Exemplo: Atualizar `hubspot_leads.csv`

```powershell
# 1. Exportar nova versão do HubSpot
# Salvar em: Downloads\hubspot_leads_novo.csv

# 2. Mover versão atual para histórico
cd "C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot"
$data = Get-Date -Format "yyyy-MM-dd"
Copy-Item hubspot_leads_ATUAL.csv "historico\hubspot_leads_$data.csv"

# 3. Substituir pela nova versão
Copy-Item "C:\Users\a483650\Downloads\hubspot_leads_novo.csv" -Destination "hubspot_leads_ATUAL.csv"

# 4. Documentar no README.md
# Adicionar linha:
## 📊 Última Atualização
- **Data**: [data atual]
- **Base**: hubspot_leads_ATUAL.csv
- **Responsável**: [seu nome]
```

✅ **Pronto!** Todos os projetos agora usam a versão atualizada automaticamente.

---

## 🚨 TROUBLESHOOTING

### Problema: "Acesso Negado" ao criar atalhos (Fase 3)

**Solução**: Execute PowerShell como Administrador

```powershell
# Clicar com botão direito no PowerShell
# "Executar como Administrador"
```

### Problema: Scripts não executam (ExecutionPolicy)

**Solução**: Liberar execução temporariamente

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

### Problema: Caminhos quebrados após reorganização

**Solução**: Atualizar caminhos nos scripts

Antes:
```python
CAMINHO_LEADS = r"C:\Users\a483650\Projetos\projeto_helio\Data\hubspot_leads.csv"
```

Depois (Opção 1 - Caminho direto):
```python
CAMINHO_LEADS = r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_ATUAL.csv"
```

Depois (Opção 2 - Via atalho):
```python
CAMINHO_LEADS = r"C:\Users\a483650\Projetos\01_Helio_ML_Producao\Data\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_ATUAL.csv"
```

### Problema: Quero desfazer tudo

**Solução**: Restaurar do backup

Se executou **FASE 3** (que cria backup automático):
```powershell
# Localizar pasta de backup (formato: _BACKUP_SEGURANCA_YYYYMMDD_HHMMSS)
cd "C:\Users\a483650\Projetos"
dir | Where-Object {$_.Name -like "*BACKUP_SEGURANCA*"}

# Restaurar arquivos necessários
```

Se fez **backup manual** antes:
- Simplesmente restaure a pasta completa do backup externo

---

## 📞 SUPORTE E DOCUMENTAÇÃO ADICIONAL

### Arquivos de Referência

| Arquivo | Localização | Conteúdo |
|---------|-------------|----------|
| **Inventário** | `Controle_de_entregas\INVENTARIO_PROJETOS_DATA_ANALYTICS.xlsx` | Catálogo de 15 projetos |
| **Duplicações** | `Controle_de_entregas\RELATORIO_DUPLICACOES.xlsx` | Análise de arquivos duplicados |
| **Relatório Fase 3** | `Controle_de_entregas\RELATORIO_REORGANIZACAO_FASE3.txt` | Detalhes da limpeza |
| **README Dados** | `_DADOS_CENTRALIZADOS\README.md` | Documentação das bases |
| **README Arquivo** | `_ARQUIVO\README.md` | Lista de projetos arquivados |

### Estrutura de Pastas Especiais

**_DADOS_CENTRALIZADOS/**: Repositório central - NUNCA deletar  
**_SCRIPTS_COMPARTILHADOS/**: Utilitários - Para futuras otimizações  
**_ARQUIVO/**: Inativos - Pode ser deletado após 6 meses  
**_BACKUP_SEGURANCA_*/**: Backup da Fase 3 - Deletar após 30 dias de validação

---

## ✅ CHECKLIST FINAL

Após completar a reorganização:

- [ ] Backup externo criado antes de começar
- [ ] Fase 1 executada com sucesso
- [ ] Fase 2 executada com sucesso
- [ ] Fase 3 executada com sucesso (ou pulada intencionalmente)
- [ ] Estrutura validada visualmente no Explorer
- [ ] Projeto Helio testado e funcionando
- [ ] Pipeline Mídia Paga testado e funcionando
- [ ] Atalhos para `_DADOS_CENTRALIZADOS` funcionando
- [ ] Inventário revisado (Excel)
- [ ] Relatório de duplicações revisado (Excel)
- [ ] README.md lido e compreendido
- [ ] Equipe notificada sobre nova estrutura
- [ ] Documentação interna atualizada
- [ ] Backup de segurança mantido por 30 dias

---

## 📅 CRONOGRAMA DE MANUTENÇÃO

### Imediato (Dia 1-7)
- Monitorar projetos em produção
- Corrigir caminhos quebrados
- Documentar problemas

### Curto Prazo (Semana 2-4)
- Validar que todos os scripts funcionam
- Ajustar processos de atualização de bases
- Treinar equipe na nova estrutura

### Médio Prazo (Mês 2)
- Deletar backup de segurança (se tudo OK)
- Extrair scripts comuns para `_SCRIPTS_COMPARTILHADOS/`
- Implementar processo de atualização automatizada

### Longo Prazo (Mês 3-6)
- Avaliar projetos em `_ARQUIVO/` para exclusão permanente
- Otimizar scripts compartilhados
- Documentar lições aprendidas

---

## 🎓 LIÇÕES APRENDIDAS

### Boas Práticas Implementadas
1. ✅ Fonte única de verdade (Single Source of Truth)
2. ✅ Versionamento de backups por data
3. ✅ Nomenclatura padronizada com prefixos numéricos
4. ✅ Separação entre Produção, Análises e Arquivo
5. ✅ Documentação centralizada

### O Que Evitar no Futuro
1. ❌ Duplicar bases entre projetos
2. ❌ Criar backups sem organização de datas
3. ❌ Nomear projetos sem padrão
4. ❌ Misturar projetos ativos com inativos
5. ❌ Falta de documentação

---

**✨ Reorganização criada em 09/02/2026**  
**📧 Dúvidas**: Consulte este README ou revise os relatórios gerados  
**🔄 Versão**: 1.0
