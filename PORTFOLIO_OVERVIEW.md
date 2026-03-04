# 📊 PORTFOLIO DE INTELIGÊNCIA & AUTOMAÇÃO DE DADOS

> **Status**: Portfólio profissional | **Nível**: Sênior em Data Science & Automação | **Conformidade**: LGPD ✓

---

## 🎯 Visão Executiva

Este repositório contém **4 projetos de produção** desenvolvidos para automação e inteligência de dados em ambiente corporativo de grande escala.

### 📈 Impacto e Números

| Métrica | Resultado |
|---------|-----------|
| **Projetos Produção** | 4 projetos P&D |
| **Modelos ML** | 2+ modelos em produção |
| **Pipelines Automáticos** | 5+ ETLs diários |
| **Integração APIs** | 3+ fontes externas (HubSpot, Meta, Google) |
| **Dados Processados** | Centenas de milhares de registros/dia |
| **Linguagem Principal** | Python 3.8+ |

---

## 📁 ESTRUTURA DO PORTFÓLIO

```
Portfolio-Dados-IA/
│
├── 📂 01_HELIO_ML_SCORING
│   ├── 🤖 Modelo de ML para Lead Scoring
│   ├── ⚙️ Feature Engineering avançado
│   ├── 🎯 Regressão & Classificação
│   └── 📊 Acurácia: ~88% (em produção)
│
├── 📂 02_PIPELINE_MIDIA_PAGA
│   ├── 🔄 ETL integrado (API Meta/Google)
│   ├── ✅ Validação automática de dados
│   ├── 📈 Cálculo de KPIs em tempo real
│   └── ⚡ Processamento de ~20k+ registros/dia
│
├── 📂 03_ANALISES_OPERACIONAIS
│   ├── 📊 Dashboards e análises ad-hoc
│   ├── 🔀 Comparativos multi-unidades
│   ├── 🎯 Eficiência de canais
│   └── 📉 Forecasting de demanda
│
├── 📂 06_FUNIL_REDBALLOON
│   ├── 🔍 Análise detalhada de conversão
│   ├── 📊 Mapeamento de funil completo
│   ├── 🔮 Cenários preditivos
│   └── 💰 Cenários de receita
│
├── 📄 README.md (este arquivo)
├── 📋 PORTFOLIO_OVERVIEW.md
├── 🚀 QUICKSTART.md
├── .gitignore (LGPD compliant)
└── 🔧 Scripts auxiliares
```

---

## 🏆 COMPETÊNCIAS DEMONSTRADAS

### 🤖 Machine Learning
- ✅ Classificação & Regressão
- ✅ Feature Engineering
- ✅ Validação e tuning de modelos
- ✅ Previsão e scoring
- ✅ Modelos em produção

### 📊 Data Engineering
- ✅ ETL Pipelines
- ✅ Integração com APIs REST
- ✅ Validação de qualidade de dados
- ✅ Processamento em batch
- ✅ Agregação e transformação de dados

### 📈 Analytics & BI
- ✅ Dashboards executivos
- ✅ KPI tracking
- ✅ Análise de funil
- ✅ Forecasting
- ✅ Análises ad-hoc

### 🔧 Engenharia & DevOps
- ✅ Automação de processos
- ✅ Scripts robustos e bem estruturados
- ✅ Tratamento de erros
- ✅ Logging e monitoramento
- ✅ Documentação técnica

---

## 🚀 COMO COMEÇAR

### 1. Clone o repositório
```bash
git clone https://github.com/seu_usuario/portfolio-dados-ia.git
cd portfolio-dados-ia
```

### 2. Explore os projetos
Cada pasta tem seu próprio `README.md` com instruções específicas:

```bash
# Projeto 1: ML Scoring
cd 01_HELIO_ML_SCORING
cat README.md

# Projeto 2: Pipeline
cd ../02_PIPELINE_MIDIA_PAGA
cat README.md

# E assim por diante...
```

### 3. Setup de um projeto
```bash
# No Windows
cd 01_HELIO_ML_SCORING
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

# No Linux/Mac
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

---

## 📂 DETALHES DE CADA PROJETO

### 🎯 **Projeto 1: HELIO ML SCORING**
`01_HELIO_ML_SCORING/`

**O que é:**
Sistema inteligente de scoring de leads para priorização automática com machine learning.

**Tecnologias:**
- Python, Pandas, Scikit-learn, XGBoost
- Feature engineering com 50+ features
- Modelo de regressão logística otimizado

**Benefício:**
Reduz tempo de análise manual de leads em 80% com scoring automático de confiabilidade.

**Como usar:**
```bash
python scripts/01_load_data.py
python scripts/02_feature_engineering.py
python scripts/03_train_model.py
python scripts/04_predict.py --input dados.csv
```

[→ Documentação completa](./01_HELIO_ML_SCORING/README.md)

---

### 🔄 **Projeto 2: PIPELINE MÍDIA PAGA**
`02_PIPELINE_MIDIA_PAGA/`

**O que é:**
Pipeline de automação end-to-end que:
- ✅ Busca dados das APIs (Meta Ads, Google Ads)
- ✅ Valida integridade e qualidade
- ✅ Calcula KPIs (ROI, CPA, ROAS)
- ✅ Gera relatórios automáticos

**Tecnologias:**
- Python, Requests, Pandas
- APIs REST (Meta, Google)
- SQL agregações

**Benefício:**
Elimina trabalho manual de relatórios. Dados atualizados em tempo real, 24/7.

**Como usar:**
```bash
# Executar pipeline completo
EXECUTAR_PIPELINE.bat

# Ou scripts individuais
python scripts/fetch_meta_ads.py
python scripts/fetch_google_ads.py
python scripts/validate_data.py
python scripts/calculate_kpis.py
```

[→ Documentação completa](./02_PIPELINE_MIDIA_PAGA/README.md)

---

### 📊 **Projeto 3: ANÁLISES OPERACIONAIS**
`03_ANALISES_OPERACIONAIS/`

**O que é:**
Suite com 5+ análises estratégicas:

| Análise | Descrição |
|---------|-----------|
| **Eficiência de Canais** | Performance de cada canal de aquisição |
| **Comparativo Unidades** | Benchmarking entre filiais |
| **Curva de Alunos** | Forecast de matriculação |
| **ICP** | Ideal Customer Profile |
| **Valor The News** | Análise de nicho específico |

**Tecnologias:**
- Pandas, Matplotlib, Plotly
- Análises exploratórias
- Modelos de previsão

**Benefício:**
Suporta decisões executivas com análises profundas e visuais atrativas.

[→ Documentação completa](./03_ANALISES_OPERACIONAIS/README.md)

---

### 🎯 **Projeto 4: FUNIL REDBALLOON**
`06_FUNIL_REDBALLOON/`

**O que é:**
Análise completa do funil de conversão incluindo:
- 🔍 Mapeamento de etapas (Leads → Clientes)
- 📊 Taxa de conversão por etapa
- 🔮 Cenários preditivos
- 💰 Projeção de receita

**Tecnologias:**
- Python, Pandas, Scipy
- Análises estatísticas
- Dashboards Excel/Power BI

**Benefício:**
Identifica gargalos no funil e oportunidades de otimização.

[→ Documentação completa](./06_FUNIL_REDBALLOON/README.md)

---

## 🔒 Segurança e Conformidade

### ✅ LGPD Compliant
- ❌ **Zero dados pessoais** em repositório
- ❌ **Zero credenciais** versionadas
- ✅ Apenas código e lógica
- ✅ `.gitignore` robusto em todas as pastas

### 📋 Como usar seu próprio dado:
1. Consulte o schema em cada projeto (`Docs/DATA_SCHEMA.md`)
2. Prepare seus dados com a mesma estrutura
3. Coloque em `data/` (ignorado pelo Git)
4. Execute os scripts normalmente

---

## 📚 Documentação

Todo projeto tem:
- 📄 `README.md` - Overview e instruções
- 📖 `Docs/` - Documentação técnica
- 📊 `scripts/` - Código limpo e documentado
- 🔧 `requirements.txt` - Dependências

---

## 🎓 Aprendizados & Tecnologias

### Linguagens
- ![Python](https://img.shields.io/badge/Python-3.8+-blue)
- ![SQL](https://img.shields.io/badge/SQL-PostgreSQL-green)
- ![Bash](https://img.shields.io/badge/Bash-Windows%20Batch-orange)

### Libraries Principais
```
pandas, numpy, scikit-learn, xgboost, requests,
matplotlib, plotly, scipy, sqlalchemy
```

### Padrões & Práticas
- ✅ Clean Code
- ✅ Versionamento Git
- ✅ Testes e validação
- ✅ Documentação técnica
- ✅ Automação e CI/CD pronto para escalar

---

## 💼 Sobre o Desenvolvedor

**Perfil:**
- Data Scientist / ML Engineer Sênior
- Especialista em automação de processos
- Experiência em startups e grandes corporações
- Forte em business intelligence e strategy

**Interesses:**
- Machine Learning aplicado
- Data Engineering e pipelines
- Otimização de processos
- Análise estratégica

---

## 📞 Contato & Links

- 🔗 [LinkedIn](https://linkedin.com) - Portfolio profissional
- 💻 GitHub - Este repositório
- 📧 Email - [seu_email@gmail.com]

---

## 📝 Licença

Código pessoal de portfólio, livre para estudo e adaptação.

---

<div align="center">

**Desenvolvido de Outubro 2025 a Março 2026 com ❤️ para demonstrar expertise em dados & inteligência**

⭐ Se achou útil, deixe uma star! ⭐

</div>
