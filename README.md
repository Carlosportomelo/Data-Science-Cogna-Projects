# 📊 Portfólio de Projetos - Inteligência de Dados & IA

> **Aviso de Segurança**: Este repositório contém **APENAS código e lógica**. Nenhum dado pessoal ou sensível foi incluído (LGPD compliant).

---

## 📁 Projetos Inclusos

### **01_Helio_ML_Producao** - Sistema de IA para Scoring de Leads
**Objetivo**: Modelo de machine learning para priorização e previsão de conversão de leads

**O que contém:**
- 🤖 Modelos ML treinados (`models/`)
- 📝 Scripts de processamento e feature engineering
- 📊 Análises e diagnósticos
- 📖 Documentação técnica

**Como usar:**
```bash
# 1. Configure o Python (requer Python 3.8+)
cd 01_Helio_ML_Producao
python -m venv .venv
.venv\Scripts\activate  # Windows

# 2. Instale dependências
pip install -r requirements.txt

# 3. Prepare seus dados (formato CSV com colunas específicas)
# Ver em Docs/ para especificação de schema

# 4. Execute o modelo
python Scripts/main_processing.py --input sua_data.csv
```

**Dependências principais:** `pandas`, `scikit-learn`, `xgboost`, `numpy`

---

### **02_Pipeline_Midia_Paga** - Automação e Análise de Performance
**Objetivo**: ETL, validação e análise de gastos em mídia paga (Meta, Google Ads)

**O que contém:**
- 🔄 Scripts de integração com APIs (Meta, Google)
- ✅ Validação de dados e qualidade
- 📈 Agregação e cálculos de KPIs
- 🐛 Ferramentas de diagnóstico

**Como usar:**
```bash
cd 02_Pipeline_Midia_Paga
python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt

# Execute o pipeline
python scripts/main_etl.py

# Ou use o batch automático
EXECUTAR_PIPELINE.bat
```

**Principais scripts:**
- `calc_sem_lead_midia_paga.py` - Calcula dias sem leads
- `verify_data.py` - Valida integridade dos dados
- `show_investimento_sem_lead_total.py` - Análise de ROI

---

### **03_Analises_Operacionais** - Análises de Negócio
**Objetivo**: Dashboards, análises de funil, eficiência de canais

**Pastas:**
- 📉 `eficiencia_canal/` - Performance de cada canal de aquisição
- 🔀 `comparativo_unidades/` - Comparação entre filiais
- 📊 `curva_alunos/` - Projeções de matriculação
- 🎯 `ICP/` - Análise de perfil ideal de cliente

**Como usar:**
```bash
cd 03_Analises_Operacionais
python scripts/gerar_dashboards.py
```

---

### **06_Analise_Funil_RedBalloon** - Análise de Conversão
**Objetivo**: Análise detalhada do funil de vendas e conversão

**O que faz:**
- 🔍 Mapeamento completo do funil
- 📊 Análise de queda em cada etapa
- 🔮 Cenários preditivos
- 📺 Dashboards interativos

**Como usar:**
```bash
cd 06_Analise_Funil_RedBalloon
python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
EXECUTAR_ANALISE.bat
```

---

## 🔐 Segurança LGPD

### O que NÃO foi incluído:
- ❌ Arquivos CSV com dados pessoais
- ❌ Excel com listas de clientes/leads
- ❌ Informações de matriculação
- ❌ Dados de HubSpot/CRM
- ❌ Credenciais, chaves de API, tokens

### O que SIM foi incluído:
- ✅ Código Python (scripts, funções, lógica)
- ✅ Modelos ML treinados (sem dados)
- ✅ Documentação e explicações
- ✅ Estrutura de projeto e arquitetura

### Como usar seu próprio dado:
1. Prepare um CSV/Excel com a **mesma estrutura de colunas**
2. Consulte o schema em `Docs/DATA_SCHEMA.md`
3. Coloque o arquivo em `data/`
4. Execute o script normalmente

---

## 🛠️ Setup Geral

### Requisitos:
- Python 3.8+
- pip ou conda
- Windows, Linux ou Mac

### Variáveis de Ambiente (se necessário):
```bash
# .env (USAR APENAS PARA DESENVOLVIMENTO LOCAL)
API_KEY_HUBSPOT=sua_chave
API_KEY_META=sua_chave
API_KEY_GOOGLE=sua_chave
```

**⚠️ Nunca commite arquivo `.env` no Git!**

---

## 📊 Arquitetura

```
cada_projeto/
├── scripts/          # Python scripts da lógica
├── Docs/            # Documentação técnica
├── models/          # Modelos treinados (se houver)
├── README.md        # Instruções específicas
├── requirements.txt # Dependências
└── .gitignore       # Arquivos a ignorar
```

---

## 🚀 Quick Start (Escolha um):

```bash
# 01_Helio_ML_Producao
cd 01_Helio_ML_Producao && python Scripts/main.py

# 02_Pipeline_Midia_Paga
cd 02_Pipeline_Midia_Paga && EXECUTAR_PIPELINE.bat

# 03_Analises_Operacionais
cd 03_Analises_Operacionais && python scripts/run_all.py

# 06_Analise_Funil_RedBalloon
cd 06_Analise_Funil_RedBalloon && EXECUTAR_ANALISE.bat
```

---

## 📝 Documentação Adicional

Veja em cada pasta:
- `README.md` - Detalhes específicos do projeto
- `Docs/` - Documentação técnica, schemas, dicionário de dados

---

## 🎓 Aprendizados Principais

Estes projetos demonstram:
- 🤖 **Machine Learning**: Modelos de scoring, regressão, classificação
- 📊 **ETL**: Integração com APIs, limpeza e validação de dados
- 📈 **Business Analytics**: KPIs, dashboards, análises de funil
- 🔧 **Automação**: Scripts para pipelines, jobs agendados
- 💾 **Big Data**: Processamento eficiente de volumes grandes

---

## 📄 Licença

Código pessoal de portfólio. Livre para uso, estudo e adaptação.

---

**Desenvolvido com ❤️ para demonstrar expertise em dados & inteligência**
