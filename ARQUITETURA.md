# 🏗️ ARQUITETURA DO PORTFÓLIO

> Visão estrutural completa do repositório. *Como os 4 projetos se encaixam em uma estratégia maior.*

---

## 🎯 Estratégia Geral

Este portfólio demonstra expertise em **3 pilares interconectados**:

```
┌─────────────────────────────────────────────────────────────┐
│                   INTELIGÊNCIA DE DADOS                      │
├─────────────────────────────────────────────────────────────┤
│                                                               │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐       │
│  │   CAPTURA    │  │ PROCESSAMENTO│  │   ANÁLISE    │       │
│  ├──────────────┤  ├──────────────┤  ├──────────────┤       │
│  │• APIs        │  │• ETL         │  │• Dashboards  │       │
│  │• Integração  │  │• Validação   │  │• Analytics   │       │
│  │• Limpeza     │  │• Scoring     │  │• Insights    │       │
│  │              │  │• Cálculos    │  │• Forecasting │       │
│  └──────────────┘  └──────────────┘  └──────────────┘       │
│                                                               │
└─────────────────────────────────────────────────────────────┘

           ↓          ↓           ↓
           
    PROJETO 02    PROJETO 01    PROJETO 03 & 04
```

---

## 📊 MAPA MENTAL DO PORTFÓLIO

```
                        ┌─────────────────────┐
                        │   DADOS BRUTOS      │
                        │  (Meta, Google,     │
                        │   HubSpot, etc)     │
                        └──────────┬──────────┘
                                   │
                    ┌──────────────┼──────────────┐
                    │              │              │
                    ▼              ▼              ▼
            ┌──────────────┐  ┌──────────────┐  └──────────────┐
            │ PROJETO 02   │  │ PROJETO 01   │  PROJETO 03/04  │
            │  PIPELINE    │  │  ML SCORING  │   ANÁLISES      │
            │   MÍDIA      │  │              │                 │
            ├──────────────┤  ├──────────────┤  ├──────────────┤
            │ Busca APIs   │  │ Carrega Data │  │ Consulta DB  │
            │ Valida       │  │ Features Eng │  │ Agrega Dados │
            │ Calcula KPIs │  │ Treina Model │  │ Gerá Análise │
            │ Armazena     │  │ Prediz Score │  │ Cria Dashboard
            └──────────┬───┘  └──────────┬───┘  └──────────┬───┘
                       │                 │                 │
                       │    ┌────────────┴────────────┐    │
                       │    │                         │    │
                       └────┼─► RELATÓRIOS/INSIGHTS ◄─────┘
                            │  & RECOMENDAÇÕES
                            └────────────────────────┘
```

---

## 🔄 FLUXO DE DADOS COMPLETO

```
1. COLETA (Projeto 02)
   ├── Meta Ads API    ──┐
   ├── Google Ads API  ──┼──► ETL Pipeline
   └── HubSpot API     ──┘
                          │
                          ▼
2. PROCESSAMENTO (Projeto 02)
   ├── Validação
   ├── Limpeza
   ├── Transformação
   └── Cálculo de KPIs ──┐
                          │
                          ├──► Armazenagem (CSV/DB)
                          │
                          └──► Projeto 01 (ML):
                                Feature Engineering
                                    │
                                    ▼
3. SCORING (Projeto 01)          Modelo Treinado
   ├── Load dados preparados        │
   ├── Apply Features               │
   ├── Predict Scores        ◄──────┘
   └── Rank Leads    ────────┐
                             │
4. ANÁLISE (Projeto 03 & 04) │
   ├── Eficiência de Canais  │
   ├── Comparativo Unidades  │
   ├── Curva de Alunos   ◄───┘
   ├── ICP
   ├── Funil & Conversão
   └── Forecasting Receita
         │
         ▼
5. SAÍDA
   ├── Dashboards
   ├── Relatórios
   ├── Alertas
   └── Recomendações
```

---

## 🏢 ESTRUTURA POR PROJETO

### HELIO ML SCORING
```
01_HELIO_ML_SCORING/
│
├── 📥 INPUT
│   └── dados.csv (leads com features)
│
├── ⚙️ PROCESSAMENTO
│   ├── Carga de dados
│   ├── Feature engineering
│   ├── Split train/test
│   └── Treinamento de modelo
│
├── 💾 MODELOS
│   ├── modelo_v1.pkl
│   ├── scaler.pkl
│   └── feature_list.json
│
├── 📊 OUTPUT
│   ├── leads_scored.csv
│   ├── model_metrics.json
│   └── predictions.parquet
│
└── 🔧 INFRAESTRUTURA
    ├── scripts/ (código)
    ├── Docs/ (documentação)
    └── requirements.txt
```

**Entrada → Processamento → Modelo → Saída**

---

### PIPELINE MÍDIA PAGA
```
02_PIPELINE_MIDIA_PAGA/
│
├── 🔌 FONTES (Externo)
│   ├── Meta Ads API
│   ├── Google Ads API
│   └── HubSpot API
│
├── 🔄 ETL PIPELINE
│   ├── Fetch dados
│   ├── Validação
│   ├── Transformação
│   ├── Agregação
│   └── Cálculo KPIs
│
├── 💾 ARMAZENAGEM
│   ├── CSV (histórico)
│   ├── Database (lookup)
│   └── Cache (performance)
│
├── 📊 OUTPUTS
│   ├── KPIs
│   ├── Relatórios
│   ├── Alertas
│   └── Dashboards
│
└── ⏰ AGENDAMENTO
    ├── Executar diário
    ├── 24/7 automático
    └── Alertas de erro
```

**APIs → ETL → Validação → Cálculo → Saída**

---

### ANÁLISES OPERACIONAIS
```
03_ANALISES_OPERACIONAIS/
│
├── 🎯 ANÁLISE 1: Eficiência de Canais
│   ├── CPL por canal
│   ├── ROI por canal
│   └── Ranking
│
├── 🎯 ANÁLISE 2: Comparativo Unidades
│   ├── Performance matriculação
│   ├── Conversão
│   └── Benchmarking
│
├── 🎯 ANÁLISE 3: Curva de Alunos
│   ├── Forecast
│   ├── Tendências
│   └── Planejamento
│
├── 🎯 ANÁLISE 4: ICP
│   ├── Perfil do melhor cliente
│   ├── Score matching
│   └── Segmentação
│
├── 🎯 ANÁLISE 5: Valor The News
│   ├── Oportunidades
│   ├── Diferenciação
│   └── Positioning
│
└── 📊 OUTPUTS
    ├── Dashboards
    ├── Relatórios
    └── Recomendações
```

**Dados → Exploração → Análise → Insights**

---

### FUNIL REDBALLOON
```
06_FUNIL_REDBALLOON/
│
├── 🔀 MAPEAMENTO FUNIL
│   ├── Etapa 1: Leads Qualificados
│   ├── Etapa 2: Propostas Enviadas
│   ├── Etapa 3: Negociação
│   ├── Etapa 4: Fechamento
│   └── Etapa 5: Receita
│
├── 📊 ANÁLISE ESTÁTICA
│   ├── Taxa de conversão por etapa
│   ├── Dropout points
│   └── Tempo médio
│
├── 🔮 ANÁLISE DINÂMICA
│   ├── Cenários "e se..."
│   ├── Impacto de mudanças
│   └── Simulações
│
└── 💰 FORECAST
    ├── Receita projetada
    ├── Cenários de crescimento
    └── Recomendações
```

**Mapeamento → Análise → Previsão → Otimização**

---

## 🔗 INTEGRAÇÃO ENTRE PROJETOS

```
┌────────────────────────────────────────────────────────────┐
│                                                              │
│  Projeto 02 (Pipeline)                                     │
│  └─► Dados limpos & validados & KPIs calculados           │
│      │                                                     │
│      ├─► Projeto 01 (ML)                                  │
│      │   └─► Leads scored & rankeados                     │
│      │       │                                             │
│      │       └──┐                                          │
│      │          │                                          │
│      └──────────┼─► Projetos 03 & 04 (Análises)          │
│                 │   ├─► Eficiência de Canais            │
│                 │   ├─► Comparativo Unidades            │
│                 │   ├─► Curva de Alunos                 │
│                 │   ├─► ICP                             │
│                 │   ├─► Funil & Conversão               │
│                 │   └─► Forecast de Receita             │
│                 │       │                                 │
│                 └───────┘                                 │
│                         │                                 │
│                         ▼                                 │
│                   DASHBOARDS                              │
│                   RELATÓRIOS                              │
│                   DECISÕES                                │
│                                                              │
└────────────────────────────────────────────────────────────┘
```

---

## 🎯 PONTOS DE SAÍDA (Outputs)

```
PROJETO 02 outputs:
├── KPI diários (ROI, CPA, etc)
├── Alertas de anomalias
└── Relatórios automáticos

PROJETO 01 outputs:
├── Score de cada lead
├── Ranking de oportunidades
└── Confiabilidade de predição

PROJETOS 03 & 04 outputs:
├── Dashboards executivos
├── Análises investigativas
├── Recomendações acionáveis
└── Forecasts de receita
```

---

## 💡 SINERGIA DOS PROJETOS

### Como se encaixam:

1. **COLETA** (P02)
   - → Traz dados "crus" da realidade
   - → Limpa e valida
   - → Calcula métricas básicas

2. **INTELIGÊNCIA** (P01)
   - ← Consome dados limpos
   - → Agrega valor com ML
   - → Score cada lead automaticamente

3. **CONTEXTO** (P03 & P04)
   - ← Consome dados scored
   - → Coloca em contexto maior
   - → Fornece insights estratégicos

---

## 🚀 Escalabilidade & Evolução

### Próximas Etapas Óbvias:

```
Fase 1 [ATUAL]
├── Scripts standalone
├── Execução manual/batch
└── CSV como armazenagem

Fase 2 [PRÓXIMA]
├── Orquestração (Airflow)
├── Database (PostgreSQL)
├── Cloud deployment
└── API endpoints

Fase 3 [FUTURA]
├── Real-time processing
├── Streaming (Kafka)
├── Advanced ML (Deep Learning)
└── MLOps infrastructure
```

---

## 📐 Padrões Técnicos Utilizados

### Em Projeto 02 (Pipeline):
- Extract-Transform-Load (ETL)
- API integration pattern
- Validation pipeline
- Logging & monitoring

### Em Projeto 01 (ML):
- Feature engineering pipeline
- Model training/inference separation
- Model persistence (pickle)
- Cross-validation

### Em Projetos 03 & 04 (Analytics):
- Exploratory Data Analysis (EDA)
- Statistical analysis
- Scenario modeling
- Dashboard generation

---

## 🎯 Por Que Esta Arquitetura?

### ✅ Separação de Responsabilidades
- Cada projeto foco em 1 coisa bem-feita
- Fácil manutenção independente
- Reutilização de componentes

### ✅ Escalabilidade
- Adicionar novo pipeline? Novo projeto
- Melhorar modelo ML? Não quebra análises
- Insights novos? Builds on top of existing

### ✅ Modularidade
- Componentes desacoplados
- Inter-operable
- Testáveis isoladamente

### ✅ Profissionalismo
- Mostra pensamento arquitetural
- Demonstra capacidade de design
- Pronto para produção

---

<div align="center">

## 📊 RESUMO ARQUITETURAL

**4 Projetos Independentes**
**3 Fluxos de Dados**
**1 Objetivo Comum: Inteligência de Negócio**

Arquitetura não é acidental, é **pensada e proposital**.

</div>

---

**Versão**: 1.0
**Data**: Março 2026 (Período: Outubro 2025 - Março 2026)
