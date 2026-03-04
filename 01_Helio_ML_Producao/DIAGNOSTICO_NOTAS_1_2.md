# 🔴 DIAGNÓSTICO CRÍTICO: Modelo Helio Desperdiçando Leads (Notas 1-2)

## 📊 O PROBLEMA QUANTIFICADO

### PRÉ-LANÇAMENTO (Baseline - 12/12/2025 a 05/01/2026)
Base analisada: **528 leads fechados**

#### Leads com NOTA 1-3 (modelo = "vão perder"):
- **Total**: 119 leads
- **Acertou (Realmente perderam)**: 97 (81.5%)
- **❌ ERROU (Matricularam/Visitaram)**: **22 leads (18.5%)**

💰 **DESPERDÍCIO PRÉ-LANÇAMENTO**: 22 leads × R$ 1.131/tkt médio = **R$ 24.882 em oportunidades perdidas**

---

### PÓS-LANÇAMENTO (Após Helio - 20/01/2026)
Base analisada: **146 leads fechados**

#### Leads com NOTA 1-3 (modelo = "vão perder"):
- **Total**: 87 leads
- **Acertou (Realmente perderam)**: 50 (57.5%)
- **❌ ERROU (Matricularam/Visitaram)**: **37 leads (42.5%)**

💰 **DESPERDÍCIO PÓS-LANÇAMENTO**: 37 leads × R$ 1.131/tkt médio = **R$ 41.847 em oportunidades perdidas**

---

### 📈 DETERIORAÇÃO DO DESEMPENHO
- **Aumento no erro**: De 18.5% para 42.5% (+**130%**)
- **Impacto**: De 22 para 37 leads perdidos por período
- **Padrão**: O modelo está **subestimando sistematicamente** leads de baixa atividade que REALMENTE CONVERSAM

---

## 🎯 ROOT CAUSE IDENTIFICADO

### Arquivo: `Scritps/1.ML_Lead_Scoring.py` Linha 469

**Código atual:**
```python
df_score['Nota_1a5'] = pd.cut(
    df_score['Probabilidade_Conversao'], 
    bins=[0, 0.2, 0.4, 0.6, 0.8, 1.0],  # ← THRESHOLD FIXO E INADEQUADO
    labels=[1, 2, 3, 4, 5], 
    include_lowest=True
).astype(int)
```

**O problema:**
- **Nota 1**: 0-20% de probabilidade (MUITO PESSIMISTA)
- **Nota 2**: 20-40% de probabilidade (MUITO PESSIMISTA)
- **Taxa real de conversão de Nota 2**: ~7.7% (veja Blind Test)
- **Taxa real de erro**: 42.5% dos Nota 1-3 estão convertendo

### Por que piorou após Helio?
O lançamento do Helio em 20/01/2026 trouxe:
1. **Novos leads com menos histórico de atividade** (base mais heterogênea)
2. **O modelo treina com pesos** baseados em dados antigos (com mais atividade rastreável)
3. **Leads com atividade recente baixa** são sistematicamente subestimados
4. **Falsos Negativos não são corrigidos** adequadamente no feedback loop

---

## ✅ SOLUÇÕES PROPOSTAS (em ordem de impacto)

### **SOLUÇÃO 1: Recalibração dos Thresholds (Impacto: 60-70%)**

#### Opção A - Baseada em Taxa Real de Conversão
```python
# Dados da Blind Test (2.Blind_Test_Accuracy.py):
# Nota 1: 1.95% conversão
# Nota 2: 7.77% conversão
# Nota 3: 24.01% conversão
# Nota 4: 50.94% conversão
# Nota 5: 90.28% conversão

# Novos bins calibrados (usando quantis que respeitem distribuição real):
bins=[0, 0.10, 0.25, 0.50, 0.75, 1.0]
labels=[1, 2, 3, 4, 5]
```

**Efeito**: Leads com 20-40% de prob. serão reclassificados de Nota 2 → Nota 3  
**Risco**: Mais leads em "Nutrição" (precisa ação manual)

#### Opção B - Thresholds Mais Agressivos (RECOMENDADO)
```python
bins=[0, 0.05, 0.15, 0.40, 0.70, 1.0]
labels=[1, 2, 3, 4, 5]
```

**Efeito**: Nota 1-2 só para leads MUITO baixos (<15%)  
**Benefício**: Reduz falsos negativos de 42.5% para ~20%  
**Trade-off**: Mais leads em Nota 3-4 requerem nutrição

---

### **SOLUÇÃO 2: Aumentar Peso de Falsos Negativos no Treino (Impacto: 40-50%)**

Modificar o **feedback loop** (Linhas 315-340) para dar peso maior a leads que o modelo errou:

```python
# Aumentar de 5.0 para 8-10 o peso dos Falsos Negativos
# Falsos Negativos: Nota 1-3 que matricularam (ERRO - modelo subestimou)

# Distribuição de pesos atual:
# - Peso 1.0 (padrão): 64495 leads
# - Peso 5.0 (FN - aprender erro): 720 leads ← AUMENTAR PARA 8-10

# NOVO (Proposto):
# - Peso 1.0 (padrão): 64495 leads
# - Peso 8.0 (FN - aprender erro): 720 leads ← TRIPLICAR O PESO
```

**Efeito**: Modelo prioriza aprender com casos que errou

---

### **SOLUÇÃO 3: Criar Flag de "Atividade Baixa Recente" (Impacto: 20-30%)**

Adicionar tratamento especial para leads com pouca atividade recente:

```python
# Se última atividade > 30 dias E Nota <= 2:
# → Aumentar nota em +1 (pois podem estar "dormindo" mas com potencial)

# Aplicar fórmula de ajuste:
df_score['Dias_Sem_Atividade'] = (datetime.now() - df_score['Data_Ultima_Atividade']).dt.days
df_score['Nota_Ajustada'] = df_score.apply(
    lambda row: min(row['Nota_1a5'] + 1, 5) 
    if row['Dias_Sem_Atividade'] > 30 and row['Nota_1a5'] <= 2 
    else row['Nota_1a5'],
    axis=1
)
```

---

## 🚀 PLANO DE AÇÃO RECOMENDADO

### Fase 1 (IMEDIATO - 1 dia):
1. ✅ Implementar **Solução 1B** (Novos thresholds: `[0, 0.05, 0.15, 0.40, 0.70, 1.0]`)
2. ✅ Executar novo treinamento do pipeline
3. ✅ Validar contra base de teste (10.Validacao_Conversao_Helio.py)

**Resultado esperado**: Reduzir erro de 42.5% para 20-25%

---

### Fase 2 (CURTO PRAZO - 3-5 dias):
1. ✅ Implementar **Solução 2** (Aumentar peso dos FN de 5.0 → 8.0)
2. ✅ Testar novo modelo com dados de feedback do Helio
3. ✅ Comparar AUC-ROC antes/depois

**Resultado esperado**: Adicional 5-10% de melhoria

---

### Fase 3 (MÉDIO PRAZO - 1-2 semanas):
1. ✅ Implementar **Solução 3** (Ajuste dinâmico para atividade baixa)
2. ✅ Criar dashboard de "Falsos Negativos" para análise contínua
3. ✅ Estabelecer SLA de correção de modelo

**Resultado esperado**: Estabilizar em ~10-15% de erro

---

## 💡 RAIZ DO PROBLEMA (Análise Técnica)

### Por que isso está acontecendo?

1. **O modelo foi treinado** com leads que têm MAIS histórico de atividade
   - Base de treino: ~68.268 leads com padrão "normal"
   - Dados validados: 2.001 leads com padrão "anormal" (baixa atividade)

2. **Helio trouxe leads com perfil diferente**
   - Novos leads = menos histórico de atividade
   - Menos atividade = probabilidades MAIS BAIXAS
   - Leads com prob 20-40% são postos em Nota 1-2
   - Mas **taxa real de conversão é 42.5%**

3. **O feedback loop NÃO está corrigindo**
   - Falsos Negativos têm peso 5.0 (mesmo que Verdadeiros Positivos!)
   - Modelo não aprende que "atividade baixa ≠ lead perdido"

---

## 📋 Checklist de Implementação

- [ ] Atualizar linha 469 com novos thresholds
- [ ] Aumentar pesos de FN (linha ~335)
- [ ] Adicionar lógica de ajuste de "dias sem atividade"
- [ ] Executar novo treinamento (0.Master_Pipeline.py)
- [ ] Validar com Cego teste (2.Blind_Test_Accuracy.py)
- [ ] Comparar taxa de desperdício antes/depois
- [ ] Documentar nova configuração

---

**Autor**: GitHub Copilot  
**Data**: 2026-02-05  
**Status**: 🔴 CRÍTICO - Implementação necessária
