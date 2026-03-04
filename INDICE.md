# 📑 ÍNDICE EXECUTIVO - Mapa Completo do Repositório

> Leia este arquivo para orientação rápida sobre tudo que está aqui

---

## 🎯 Comece Por Aqui (5 min)

```
1. PORTFOLIO_OVERVIEW.md ← Você está aqui! Visão geral do portfólio
2. README.md ← Como começar
3. Escolha um projeto abaixo
```

---

## 📂 Estrutura do Repositório

```
portfolio-dados-ia/
│
├── 📋 DOCUMENTAÇÃO GERAL
│   ├── README.md                    ← Documentação Principal (START HERE)
│   ├── PORTFOLIO_OVERVIEW.md        ← Este arquivo - Visão geral
│   ├── QUICKSTART.md                ← Como começar em 5 min
│   ├── CONQUISTAS_IMPACTO.md        ← Resultados & impacto
│   ├── COMO_FAZER_PUSH.md           ← Tutorial GitHub
│   └── INDEX.md                     ← Este mapa (você lê agora)
│
├── 🚀 SCRIPTS AUXILIARES
│   ├── preparar_github.py           ← Remove dados sensíveis
│   ├── gerar_requirements.py        ← Gera requirements.txt
│   └── validar_seguranca_github.py  ← Verifica LGPD compliance
│
├── 📂 HELIO ML SCORING
│   ├── README.md                    ← Como usar este projeto
│   ├── Docs/
│   │   ├── DATA_SCHEMA.md          ← Estrutura esperada dos dados
│   │   ├── FEATURES.md             ← Features utilizadas
│   │   └── MODEL_PERFORMANCE.md    ← Métricas do modelo
│   ├── scripts/
│   │   ├── 01_load_data.py
│   │   ├── 02_feature_engineering.py
│   │   ├── 03_train_model.py
│   │   └── 04_predict.py
│   ├── models/ (treinados + prontos)
│   └── requirements.txt
│
├── 📂 PIPELINE MÍDIA PAGA
│   ├── README.md
│   ├── Docs/
│   │   ├── API_INTEGRATION.md      ← Como integrar com APIs
│   │   ├── KPI_DEFINITIONS.md      ← Definição de KPIs
│   │   └── PIPELINE_FLOW.md        ← Fluxo do ETL
│   ├── scripts/
│   │   ├── fetch_meta_ads.py
│   │   ├── fetch_google_ads.py
│   │   ├── validate_data.py
│   │   ├── calculate_kpis.py
│   │   └── generate_reports.py
│   ├── EXECUTAR_PIPELINE.bat       ← Rodar tudo em 1 clique
│   └── requirements.txt
│
├── 📂 ANÁLISES OPERACIONAIS
│   ├── README.md
│   ├── Docs/
│   │   └── ANALISES_DISPONÍVEIS.md ← O que tem de analysis
│   ├── eficiencia_canal/
│   │   ├── README.md
│   │   └── scripts/
│   ├── comparativo_unidades/
│   │   ├── README.md
│   │   └── scripts/
│   ├── curva_alunos/
│   │   ├── README.md
│   │   └── scripts/
│   ├── ICP/
│   │   ├── README.md
│   │   └── scripts/
│   └── valor_the_news/
│       ├── README.md
│       └── scripts/
│
├── 📂 FUNIL REDBALLOON
│   ├── README.md
│   ├── Docs/
│   │   ├── FUNIL_MAPEAMENTO.md     ← Etapas do funil
│   │   ├── CENARIOS_PREDITIVOS.md  ← Análise de cenários
│   │   └── RECEITA_FORECAST.md     ← Projeção de receita
│   ├── scripts/
│   │   ├── map_funnel.py
│   │   ├── calculate_dropoff.py
│   │   ├── simulate_scenarios.py
│   │   └── forecast_revenue.py
│   ├── EXECUTAR_ANALISE.bat        ← Rodar análise
│   └── requirements.txt
│
├── .gitignore                      ← Segurança LGPD
└── [data/, outputs/]               ← Pastas ignoradas por Git
```

---

## 🎓 NAVEGAÇÃO POR CASO DE USO

### Sua Profissão/Interesse → Onde Começar

| Perfil | Comece Por | Estude | Aprenda |
|--------|-----------|--------|---------|
| **Data Scientist** | [01_HELIO_ML](./01_HELIO_ML_SCORING/README.md) | Feature Engineering | ML em produção |
| **Data Engineer** | [02_PIPELINE](./02_PIPELINE_MIDIA_PAGA/README.md) | ETL & APIs | Automação |
| **Líder/Executivo** | [03_ANALISES](./03_ANALISES_OPERACIONAIS/README.md) | Dashboards | Business Analytics |
| **Gerente de Vendas** | [06_FUNIL](./06_FUNIL_REDBALLOON/README.md) | Conversão | Previsibilidade |
| **Curiosidade Geral** | [QUICKSTART.md](./QUICKSTART.md) | Todos | Tudo um pouco |

---

## 📖 DOCUMENTAÇÃO DISPONÍVEL

### Arquivos Principais (Este Repo)
| Arquivo | Tempo Leitura | Objetivo |
|---------|---------------|----------|
| [README.md](./README.md) | 10 min | Overview & setup |
| [PORTFOLIO_OVERVIEW.md](./PORTFOLIO_OVERVIEW.md) | 15 min | Visão completa |
| [CONQUISTAS_IMPACTO.md](./CONQUISTAS_IMPACTO.md) | 20 min | Resultados |
| [QUICKSTART.md](./QUICKSTART.md) | 5 min | Start rápido |
| [COMO_FAZER_PUSH.md](./COMO_FAZER_PUSH.md) | 15 min | GitHub setup |

### Documentação Por Projeto
- **01_HELIO_ML_SCORING/**
  - `Docs/DATA_SCHEMA.md` - Formato de entrada esperado
  - `Docs/FEATURES.md` - Explicação das 50+ features
  - `Docs/MODEL_PERFORMANCE.md` - Métricas e validação

- **02_PIPELINE_MIDIA_PAGA/**
  - `Docs/API_INTEGRATION.md` - Como usar as APIs
  - `Docs/KPI_DEFINITIONS.md` - Como cada KPI é calculado
  - `Docs/PIPELINE_FLOW.md` - Fluxograma do ETL

- **03_ANALISES_OPERACIONAIS/**
  - Cada subpasta tem seu próprio README.md
  - Explicação de cada análise disponível

- **06_FUNIL_REDBALLOON/**
  - `Docs/FUNIL_MAPEAMENTO.md` - Etapas e definições
  - `Docs/CENARIOS_PREDITIVOS.md` - Como rodar cenários
  - `Docs/RECEITA_FORECAST.md` - Projeção de receita

---

## 🚀 COMO EXECUTAR CADA PROJETO

### Setup Rápido (todos igual)
```bash
# 1. Clone
git clone <seu_repo_url>
cd seu_repo

# 2. Entre no projeto desejado
cd 01_HELIO_ML_SCORING  # Escolha um

# 3. Setup venv
python -m venv .venv
.venv\Scripts\activate

# 4. Instale dependências
pip install -r requirements.txt

# 5. Execute!
python scripts/main.py  # Ou scripts específicos
```

### Execução Detalhada Por Projeto

**HELIO ML SCORING:** [→ Ir para README](./01_HELIO_ML_SCORING/README.md)
```bash
python scripts/01_load_data.py
python scripts/02_feature_engineering.py
python scripts/03_train_model.py
python scripts/04_predict.py --input seu_dado.csv
```

**PIPELINE MÍDIA PAGA:** [→ Ir para README](./02_PIPELINE_MIDIA_PAGA/README.md)
```bash
# Automático
EXECUTAR_PIPELINE.bat

# Ou manual
python scripts/fetch_meta_ads.py
python scripts/calculate_kpis.py
```

**ANÁLISES OPERACIONAIS:** [→ Ir para README](./03_ANALISES_OPERACIONAIS/README.md)
```bash
cd eficiencia_canal
python scripts/run_analysis.py
```

**FUNIL REDBALLOON:** [→ Ir para README](./06_FUNIL_REDBALLOON/README.md)
```bash
EXECUTAR_ANALISE.bat
```

---

## 🔒 SEGURANÇA & CONFORMIDADE

### ✅ LGPD Compliant
- ❌ Nenhum CSV com dados pessoais
- ❌ Nenhuma credencial
- ✅ Apenas código
- ✅ `.gitignore` configurado

**Como adicionar seus dados:**
1. Prepare CSV com mesmo schema que em `Docs/DATA_SCHEMA.md`
2. Coloque em `data/` (será ignorado por Git)
3. Execute scripts normahlmente

---

## 📊 ESTATÍSTICAS DO REPOSITÓRIO

```
Projetos:        4 completos
Scripts Python:  20+
Documentação:    10+ arquivos
Modelos ML:      2+ prontos
Linhas de Código: 5000+
Tamanho:         ~10MB (sem dados)
```

---

## 🎯 ROTEIROS RECOMENDADOS

### Para Aprender (1 hora)
1. Leia [CONQUISTAS_IMPACTO.md](./CONQUISTAS_IMPACTO.md) (15 min)
2. Execute [QUICKSTART.md](./QUICKSTART.md) (20 min)
3. Explore código de 1 projeto (25 min)

### Para Reproduzir (2 horas)
1. Setup venv (5 min)
2. Prepare dados (30 min)
3. Execute projeto (30 min)
4. Adapte para seus dados (55 min)

### Para Entender Profundo (4 horas)
1. Leia documentação principal (1 h)
2. Explore código com IDE (1 h)
3. Debug & test (1 h)
4. Extensões & melhorias (1 h)

---

## 🔍 PROCURA POR ALGO ESPECÍFICO?

| Procura | Arquivo |
|---------|---------|
| Feature Engineering | `01_HELIO_ML_SCORING/Docs/FEATURES.md` |
| Integração API | `02_PIPELINE_MIDIA_PAGA/Docs/API_INTEGRATION.md` |
| Cálculo KPI | `02_PIPELINE_MIDIA_PAGA/Docs/KPI_DEFINITIONS.md` |
| Análises Disponíveis | `03_ANALISES_OPERACIONAIS/Docs/` |
| Funil de Vendas | `06_FUNIL_REDBALLOON/Docs/FUNIL_MAPEAMENTO.md` |
| Como Fazer Push | `COMO_FAZER_PUSH.md` |
| Quick Start | `QUICKSTART.md` |
| Impacto & Resultados | `CONQUISTAS_IMPACTO.md` |

---

## 💼 INFORMAÇÕES DO AUTOR

**Experiência**: Data Science & Automação Corporativa
**Expertise**: ML, ETL, Analytics, Business Intelligence
**Linguagem Principal**: Python
**Nível**: Intermediário-Avançado

---

## 📝 LICENÇA & USO

**Tipo**: Portfólio pessoal
**Licença**: Livre para uso, estudo e adaptação
**Restrição**: Não inclua dados pessoais reais

---

<div align="center">

## 🎬 PRONTO PARA COMEÇAR?

**Próximo Passo →** Escolha um arquivo acima ou project

Tempo estimado para primeira execução: **10 minutos**

Questions? Consulte o README.md do projeto específico

</div>

---

**Last Updated**: Março 2026 (Período: Out 2025 - Mar 2026)
**Versão Este Índice**: 1.0
