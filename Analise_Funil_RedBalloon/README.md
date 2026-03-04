# 🎈 Análise de Funil - Red Balloon

Relatório completo de performance do funil de vendas da **Red Balloon**, com foco nos **ciclos de captação** (Outubro → Março), comparando performance ciclo vs ciclo.

---

## 📅 Ciclos de Captação

A análise é focada nos **períodos de alta captação**:

- **Ciclo 22.1**: Out/2021 → Mar/2022
- **Ciclo 23.1**: Out/2022 → Mar/2023
- **Ciclo 24.1**: Out/2023 → Mar/2024
- **Ciclo 25.1**: Out/2024 → Mar/2025
- **Ciclo 26.1**: Out/2025 → Mar/2026

⚠️ **Leads de Abril a Setembro NÃO são incluídos** (fora do ciclo de alta)

---

## 🚀 Execução Rápida

**Duplo clique no arquivo:**
```
EXECUTAR_ANALISE.bat
```

Ou via terminal:
```powershell
.\EXECUTAR_ANALISE.bat
```

---

## 📁 Estrutura do Projeto

```
06_Analise_Funil_RedBalloon/
├── EXECUTAR_ANALISE.bat              # ⭐ Execute este arquivo!
├── data/                              # Dados processados
│   └── hubspot_leads_redballoon.csv  # Base filtrada (gerada automaticamente)
├── outputs/                           # Relatórios gerados
│   └── funil_redballoon_[data].xlsx  # Relatório principal
└── scripts/
    ├── 0.preparar_dados.py           # Filtra dados do HubSpot
    └── 1.analise_funil_redballoon.py # Gera análise completa
```

---

## 📊 Relatório Gerado

O arquivo Excel contém **8 abas** com análises focadas em números:

### 1️⃣ Resumo Por Ciclo
- Total de leads por ciclo (22.1, 23.1, 24.1, 25.1, 26.1)
- Quantidade de matriculados, perdidos e em qualificação
- **Taxa de conversão por ciclo** (%)
- **Taxa de perda** (%)
- **Comparativo ciclo vs ciclo**

### 2️⃣ Status Atual
- Distribuição atual de todos os leads
- Quantidade e percentual por status

### 3️⃣ Performance Unidade
- Total de leads por unidade
- Matriculados, perdidos e em qualificação por escola
- Taxa de conversão por unidade

### 4️⃣ Performance Fonte
- Total de leads por fonte (Meta Ads, Google Ads, Offline, etc.)
- Conversão por origem do lead
- Identificação das melhores fontes

### 5️⃣ Evolução Mensal
- Total de leads mês a mês
- Status dos leads por período
- Taxa de conversão mensal
Ciclo x Unidade
- Cruzamento de ciclos com unidades
- Identificar crescimento/queda por escola ao longo dos ciclos

### 7️⃣ Matriz Ciclo x Fonte
- Cruzamento de ciclos com fontes
- Evolução de cada canal nos períodos de captaçã
- Evolução de cada canal ao longo do tempo

### 8️⃣ Em Qualificação
- Lista detalhada de todos os leads ativos
- Informações de contato e histórico
- Leads que ainda podem converter

---

## 🔄 Fluxo de Trabalho

### Atualização da Base

1. **Sincronizar dados do HubSpot:**
   ```powershell
   python ..\_SCRIPTS_COMPARTILHADOS\sincronizar_bases.py
   ```

2. **Executar análise:**
   ```powershell
   .\EXECUTAR_ANALISE.bat
   ```

3. **Conferir resultado:**
   - Abrir arquivo em `outputs/funil_redballoon_[data].xlsx`

---

## 📈 Métricas Principais

### Taxa de Conversão
```
Taxa de Conversão = (Matriculados / Total de Leads) × 100
```

### Taxa de Perda
```
Taxa de Perda = (Perdidos / Total de Leads) × 100
```

### Leads Ativos
```
Em Qualificação = Leads que ainda não foram finalizados
```

---

## 🎯 Interpretação dos Dados

### Leads Matriculados
✅ Leads que completaram o processo e se matricularam

### Leads Perdidos  
❌ Leads que foram descartados ou não converteram

### Leads em Qualificação
⏳ Leads ativos que ainda estão sendo trabalhados

---

## 🔍 Filtros Aplicados

O script filtra automaticamente apenas leads do pipeline:
- `Red Balloon - Unidades de Rua`
- Qualquer etapa que contenha "Red Balloon" no nome
- **Apenas meses de Outubro a Março** (ciclos de captação)

---

## 🎯 Classificação de Qualidade dos Leads

### 🔥 Muito Quente
- 5+ atividades + contato nos últimos 7 dias
- **Ação:** Prioridade máxima - fechar agora

### 🟠 Quente
- 5+ atividades + contato em até 30 dias, OU
- 2-4 atividades + contato nos últimos 7 dias
- **Ação:** Alta prioridade - engajar esta semana

### 🟡 Morno / Esfriando
- 2-4 atividades com algum tempo sem contato, OU
- 1 atividade com contato recente
- **Ação:** Reaquecer - novo contato urgente

### ❄️ Frio / Muito Frio
- Poucas atividades + muito tempo sem contato, OU
- Sem atividades registradas
- **Ação:** Reativar - campanhas de reengajamento

---

## 📊 Exemplos de Insights

Com este relatório você pode responder:

- **Quantos leads geramos em cada ciclo?** → Aba 1
- **Qual ciclo teve melhor taxa de conversão?** → Aba 1
- **Há tendência de crescimento ou queda entre ciclos?** → Aba 1, 5
- **Quais unidades têm melhor performance?** → Aba 3, 6
- **Qual fonte traz mais leads nos ciclos de alta?** → Aba 4, 7
- **Qual fonte tem melhor taxa de conversão?** → Aba 4
- **Quantos leads temos ativos agora?** → Aba 2
- **Como está a performance do ciclo atual (26.1)?** → Aba 1
- **Houve crescimento ou queda por unidade entre ciclos?** → Aba 6
- **Qual a qualidade dos leads em qualificação?** → Aba 11 (coluna Qualidade_Lead) 🆕
- **Quanto tempo os leads existem?** → Aba 8, 11 (coluna Dias_Desde_Criacao) 🆕
- **Há quanto tempo não contatamos cada lead?** → Aba 11 (coluna Dias_Sem_Contato) 🆕
- **Quais leads tiveram visita presencial?** → Aba 11 (coluna Teve_Visita) 🆕
- **Quantos leads precisamos gerar para bater a meta?** → Aba 10 🆕
- **Os leads atuais são suficientes?** → Aba 10 🆕

---

## ⚙️ Configurações

### Fontes Classificadas

- **Meta Ads**: Facebook/Instagram (social pago)
- **Google Ads**: Pesquisa paga, Google
- **SEO/Orgânico**: Pesquisa orgânica
- **Offline**: Fontes off-line, CRM direto
- **Direto**: Acesso direto ao site
- **Outros**: Demais fontes

### Status dos Leads

- **Matriculado**: Etapa contém "MATRÍCULA" ou "CONCLUÍDA"
- **Perdido**: Etapa contém "PERDIDO"
- **Em Qualificação**: Etapa contém "QUALIFICAÇÃO"

---

## 🐛 Troubleshooting

### Erro: "Arquivo não encontrado"
**Solução:**
```powershell
python ..\_SCRIPTS_COMPARTILHADOS\sincronizar_bases.py
```

### Nenhum lead encontrado
**Possíveis causas:**
- Base do HubSpot não contém dados da Red Balloon
- Nome do pipeline mudou
- Dados ainda não foram sincronizados

### Arquivo Excel não abre
**Solução:**
- Fechar o Excel se estiver com o arquivo aberto
- Executar novamente o script

---

## 📝 Dados de Origem

**Fonte primária:** HubSpot CRM  
**Base utilizada:** `_DADOS_CENTRALIZADOS/hubspot/hubspot_leads_ATUAL.csv`  
**Pipeline:** Red Balloon - Unidades de Rua

---

## 🔄 Frequência de Atualização

**Recomendado:**
- Mensal (início de cada mês)
- Ou sempre que precisar de números atualizados

**Tempo de execução:** < 30 segundos

---

## 📞 Suporte

Para dúvidas sobre:
- **Dados**: Verificar base do HubSpot
- **Métricas**: Validar classificação de status
- *3.0 - Fevereiro 2026** 🆕
- ✅ Novas abas: Qualidade, Engajamento, Projeção de Meta
- ✅ Classificação de qualidade dos leads (Quente/Morno/Frio)
- ✅ Cálculo de idade do lead (dias desde criação)
- ✅ Identificação de último contato (dias sem contato)
- ✅ Indicador de visita presencial (3+ atividades)
- ✅ Simulação de metas (precisa gerar novos leads?)
- ✅ Taxa de conversão histórica automática

**v*Performance**: Analisar filtros aplicados

---

## 📝 Changelog

**v2.0 - Fevereiro 2026**
- ✅ Análise por ciclos de captação (Out-Mar)
- ✅ Comparativo ciclo vs ciclo
- ✅ Filtro automático de período de alta
- ✅ 5 ciclos identificados (22.1 até 26.1)

**v1.0 - Fevereiro 2026**
- ✅ Análise completa de funil
- ✅ 8 abas de relatório
- ✅ Métricas de conversão
- ✅ Performance por unidade e fonte
- ✅ Evolução temporal
- ✅ Leads em qualificação detalhados
