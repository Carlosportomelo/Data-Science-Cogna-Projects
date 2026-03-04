# 📊 Pipeline de Análise de Performance de Mídia Paga

Pipeline integrado para análise de performance de campanhas de mídia paga, consolidando dados de **Meta Ads**, **Google Ads** e **HubSpot CRM**. 

O sistema realiza limpeza de dados, normalização de funil de vendas, atribuição de investimento por lead (prorrateio) e gera relatórios completos para análise de RVO, Matrículas por Ciclo de Captação e Cenários Preditivos.

---

## 🚀 Execução Rápida

### Opção 1: Executar Todo o Pipeline (RECOMENDADO)

Clique 2x no arquivo:
```
EXECUTAR_PIPELINE.bat
```

Ou via linha de comando:
```powershell
# Windows (Batch)
.\EXECUTAR_PIPELINE.bat

# Python (mais detalhado)
python scripts\0.executar_pipeline_completo.py
```

Isso executa automaticamente todas as 4 etapas em sequência:
1. ✅ Análise Meta Ads
2. ✅ Análise Google Ads
3. ✅ Integração HubSpot
4. ✅ Cenários Preditivos

---

## 📁 Estrutura do Projeto

```
02_Pipeline_Midia_Paga/
├── EXECUTAR_PIPELINE.bat              # ⭐ Execute este arquivo!
├── data/                              # Dados de entrada
│   ├── meta_dataset.csv              # Base Meta Ads
│   ├── googleads_dataset.csv         # Base Google Ads
│   └── hubspot_leads.csv             # Base HubSpot CRM
├── outputs/                           # Resultados gerados
│   ├── meta_dataset_dashboard.xlsx
│   ├── google_dashboard.xlsx
│   ├── meta_googleads_blend_[data].xlsx
│   └── cenarios_preditivos_[data].xlsx
└── scripts/                           # Scripts do pipeline
    ├── 0.executar_pipeline_completo.py   # Script master (Python)
    ├── 1.analise_performance_meta.py     # Processa Meta Ads
    ├── 2.analise_performance_google.py   # Processa Google Ads
    ├── 3.analise_performance_hubspot_FINAL_ID.py  # Integração HubSpot
    └── 4.gerar_cenarios_preditivos.py    # Cenários e projeções
```

---

## 🔄 Fluxo de Trabalho

### 1️⃣ Atualizar Bases de Dados

**Baixe os dados mais recentes** (últimos 30 dias) das plataformas:

#### Meta Ads
1. Acesse o Meta Business Suite
2. Exporte dados de campanhas (últimos 30 dias)
3. Salve como: `data/meta_dataset.csv`

#### Google Ads  
1. Acesse o Google Ads Manager
2. Exporte relatório de performance (últimos 30 dias)
3. Salve como: `data/googleads_dataset.csv`

#### HubSpot (Opcional - sincronização automática)
- Execute: `python ..\_SCRIPTS_COMPARTILHADOS\sincronizar_bases.py`
- Ou substitua manualmente: `data/hubspot_leads.csv`

### 2️⃣ Executar Pipeline

```powershell
.\EXECUTAR_PIPELINE.bat
```

✅ O pipeline vai:
- Carregar histórico completo existente
- Mesclar com os novos dados (últimos 30 dias)
- **Preservar todo o histórico anterior**
- Gerar todos os relatórios atualizados

### 3️⃣ Conferir Resultados

Os arquivos são salvos em `outputs/`:
- `meta_dataset_dashboard.xlsx` - Dashboard Meta Ads
- `google_dashboard.xlsx` - Dashboard Google Ads
- `meta_googleads_blend_[data].xlsx` - Base consolidada
- `cenarios_preditivos_[data].xlsx` - Projeções

---

## 📋 Detalhamento dos Scripts

### 1. analise_performance_meta.py
**Função**: Processa dados do Meta Ads (Facebook/Instagram)

**Input**: `data/meta_dataset.csv`

**Output**: `outputs/meta_dataset_dashboard.xlsx` com abas:
- Meta_Completo (histórico completo)
- Meta_YoY (análise ano a ano)
- Meta_2023, Meta_2024, Meta_2025

**Features**:
- ✅ Lógica incremental (preserva histórico)
- ✅ Normalização de funil
- ✅ Cálculo de KPIs (CPL, CTR, CPC, etc.)

### 2. analise_performance_google.py
**Função**: Processa dados do Google Ads

**Input**: `data/googleads_dataset.csv` (ignora 2 primeiras linhas)

**Output**: `outputs/google_dashboard.xlsx` com abas:
- Google_Completo (histórico completo)
- Google_YoY (análise ano a ano)
- Google_2023, Google_2024, Google_2025

**Features**:
- ✅ Lógica incremental (preserva histórico)
- ✅ Conversão de formatos monetários
- ✅ KPIs padrão do Google Ads

### 3. analise_performance_hubspot_FINAL_ID.py
**Função**: Integra leads do HubSpot com investimentos de Meta/Google

**Input**: 
- `data/hubspot_leads.csv`
- `outputs/meta_dataset_dashboard.xlsx`
- `outputs/google_dashboard.xlsx`

**Output**: `outputs/meta_googleads_blend_[timestamp].xlsx`

**Features**:
- ✅ Prorrateio de investimento por lead
- ✅ Normalização de funil (Lead → RVO → Matrícula)
- ✅ Atribuição por ciclo de captação
- ✅ Análise por produto/unidade

### 4. gerar_cenarios_preditivos.py
**Função**: Gera projeções e análise comparativa

**Input**: `outputs/meta_googleads_blend_*.xlsx` (mais recente)

**Output**: `outputs/cenarios_preditivos_[timestamp].xlsx`

**Features**:
- ✅ Análise Realizado vs Simulado
- ✅ Projeções para 12 meses
- ✅ Comparativo de estratégias

---

## 🔐 Segurança de Dados

### Histórico Preservado
✅ Os scripts **NUNCA apagam o histórico**

Quando você baixa apenas 30 dias:
1. Script carrega histórico completo
2. Identifica data mínima dos novos dados
3. Mantém todo histórico anterior a essa data
4. Substitui apenas os últimos 30 dias

**Exemplo**:
- Histórico: 01/01/2023 → 09/02/2026
- Baixa: 12/01/2026 → 11/02/2026  
- Resultado: 01/01/2023 → 11/02/2026 ✅

### Backup Automático
Os arquivos antigos são preservados com timestamp no nome.

---

## ⚙️ Manutenção

### Sincronização de Bases Centrais
```powershell
python ..\\_SCRIPTS_COMPARTILHADOS\\sincronizar_bases.py
```

### Verificar Versões das Bases
```powershell
dir data\*.csv
```

### Limpar Outputs Antigos (opcional)
```powershell
# Manter apenas os 5 mais recentes
Get-ChildItem outputs\meta_googleads_blend_*.xlsx | 
  Sort-Object LastWriteTime -Descending | 
  Select-Object -Skip 5 | 
  Remove-Item
```

---

## 🐛 Troubleshooting

### Erro: "Arquivo não encontrado"
- Verifique se as bases estão em `data/`
- Confirme os nomes dos arquivos

### Erro: "Nenhum histórico encontrado"
- Normal na primeira execução
- O script criará os arquivos automaticamente

### Erro na Etapa 3 (HubSpot)
- Verifique se as etapas 1 e 2 geraram os dashboards
- Confira se `hubspot_leads.csv` está atualizado

---

## 📊 Integração com BI

### Looker Studio / Power BI
Conecte aos arquivos:
- `outputs/meta_googleads_blend_[data].xlsx` (base principal)
- `outputs/cenarios_preditivos_[data].xlsx` (projeções)

### Atualização
1. Baixe dados das plataformas
2. Execute `EXECUTAR_PIPELINE.bat`
3. Atualize as fontes no BI

---

## 📝 Changelog

**v2.0 - Fevereiro 2026**
- ✅ Criado script de execução automática (`EXECUTAR_PIPELINE.bat`)
- ✅ Adicionado `0.executar_pipeline_completo.py`
- ✅ Integração com repositório central de dados
- ✅ Documentação completa

**v1.0 - Janeiro 2026**  
- ✅ Pipeline manual com 4 scripts independentes
- ✅ Lógica incremental para preservação de histórico
