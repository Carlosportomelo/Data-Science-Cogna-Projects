# ✅ GUIA DE USO - ECOSSISTEMA DE DATA SCIENCE

**Última Atualização**: 10/02/2026  
**Status**: ✅ Sistema Operacional  
**Período de Desenvolvimento**: Outubro 2025 - Fevereiro 2026

---

## 🎉 SISTEMA IMPLEMENTADO

Este ecossistema de Data Science foi desenvolvido do zero entre outubro de 2025 e fevereiro de 2026, seguindo best practices de engenharia de dados e governança. Toda a infraestrutura está operacional e documentada.

### ✅ Componentes Implementados

**🏗️ Infraestrutura Central**
- ✅ `_DADOS_CENTRALIZADOS/` com 5 bases oficiais (41 MB)
- ✅ `_SCRIPTS_COMPARTILHADOS/` com 4 utilitários Python
- ✅ `_ARQUIVO/` com 5 projetos inativos preservados

**🚀 Projetos em Produção**
- ✅ `01_Helio_ML_Producao/` - Lead Scoring (Random Forest)
- ✅ `02_Pipeline_Midia_Paga/` - ROI Meta Ads

**📊 Projetos de Análises**
- ✅ `03_Analises_Operacionais/` - 4 projetos consolidados
- ✅ `04_Auditorias_Qualidade/` - 2 projetos de validação
- ✅ `05_Pesquisas_Educacionais/` - 2 estudos de perfil

---

## 📁 ESTRUTURA COMPLETA

```
C:\Users\a483650\Projetos\
│
├── 📂 _DADOS_CENTRALIZADOS/          ← Fonte única oficial (SSOT)
│   ├── hubspot/
│   │   ├── hubspot_leads_ATUAL.csv (29.37 MB)
│   │   ├── hubspot_negocios_perdidos_ATUAL.csv (10.85 MB)
│   │   └── historico/
│   ├── matriculas/
│   │   ├── matriculas_finais_ATUAL.csv (0.16 MB)
│   │   ├── matriculas_finais_ATUAL.xlsx (0.79 MB)
│   │   └── historico/
│   └── marketing/
│       ├── meta_ads_ATUAL.csv (0.15 MB)
│       └── historico/
│
├── 📂 _SCRIPTS_COMPARTILHADOS/       ← Automações reutilizáveis
│   ├── sincronizar_bases.py         ← Sincroniza bases automaticamente
│   ├── validar_reorganizacao.py     ← Valida integridade do sistema
│   ├── inventario_projetos.py       ← Gera inventário Excel
│   └── analisar_duplicacoes.py      ← Detecta duplicatas MD5
│
├── 📂 _ARQUIVO/                      ← Projetos descontinuados (5)
│
├── 📂 01_Helio_ML_Producao/          ← PRODUÇÃO ⭐
│   ├── Scritps/
│   │   └── 1.ML_Lead_Scoring.py
│   ├── Data/ (sincronizado automaticamente)
│   └── Outputs/
│       ├── Dados_Scored/
│       ├── Relatorios_Unidades/
│       └── Relatorios_ML/
│
├── 📂 02_Pipeline_Midia_Paga/        ← PRODUÇÃO ⭐
│   ├── scripts/
│   │   └── pipeline_completo.py
│   └── data/ (sincronizado automaticamente)
│
├── 📂 03_Analises_Operacionais/      ← 4 projetos
│   ├── eficiencia_canal/
│   ├── comparativo_unidades/
│   ├── curva_alunos/
│   └── valor_the_news/
│
├── 📂 04_Auditorias_Qualidade/       ← 2 projetos
│   ├── auditoria_leads_sumidos/
│   └── correcao_meta_callcenter/
│
├── 📂 05_Pesquisas_Educacionais/     ← 2 projetos
│   ├── perfil_socioeconomico/
│   └── analises_educacionais/
│
└── 📂 Controle_de_entregas/          ← Documentação Master
    ├── README.md
    ├── PROJETO_DATA_SCIENCE.md
    ├── ARQUITETURA_ECOSSISTEMA.md
    ├── GUIA_USO.md (este arquivo)
    ├── SCRIPT_APRESENTACAO.md
    ├── RESUMO_EXECUTIVO.md
    └── [demais documentos...]
```

---

## 🚀 COMO USAR O ECOSSISTEMA

### 📥 1. Atualizar Bases Centralizadas

Quando receber novos dados oficiais (HubSpot, matrículas, etc.):

```bash
# Passo 1: Navegue até o diretório da base
cd C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot

# Passo 2: Mova a versão atual para histórico (com data)
Move-Item hubspot_leads_ATUAL.csv historico/hubspot_leads_2026-02-10.csv

# Passo 3: Coloque o novo arquivo e renomeie para _ATUAL
Copy-Item C:\Downloads\novo_hubspot.csv hubspot_leads_ATUAL.csv

# Passo 4: Execute sincronização automática
cd C:\Users\a483650\Projetos\_SCRIPTS_COMPARTILHADOS
python sincronizar_bases.py
```

**O que o script de sincronização faz:**
- Copia automaticamente as bases _ATUAL para todos os projetos que as utilizam
- Valida integridade dos arquivos
- Gera log de sincronização
- Garante que todos os projetos usam a versão oficial

---

### 🤖 2. Executar Projetos de Produção

#### Lead Scoring (Helio ML)
```bash
# Navegue até o projeto
cd C:\Users\a483650\Projetos\01_Helio_ML_Producao\Scritps

# Execute o script principal
python 1.ML_Lead_Scoring.py

# Outputs gerados em:
# ../Outputs/Dados_Scored/
# ../Outputs/Relatorios_Unidades/
# ../Outputs/Relatorios_ML/
```

**O que o script faz:**
- Carrega dados do HubSpot + Matrículas (40 MB)
- Feature engineering (30+ features)
- Treina modelo Random Forest
- Gera relatórios automatizados por unidade
- Salva leads scorados com probabilidade de conversão

#### Pipeline Meta Ads
```bash
cd C:\Users\a483650\Projetos\02_Pipeline_Midia_Paga\scripts
python pipeline_completo.py

# Outputs gerados em ../data/outputs/
```

**O que o script faz:**
- Integra dados Meta Ads + HubSpot
- Calcula métricas: ROI, CAC, LTV, conversão
- Gera dashboards automatizados
- Relatórios de performance de campanhas

---

### 🔍 3. Validar Integridade do Sistema

Após qualquer modificação estrutural:

```bash
cd C:\Users\a483650\Projetos\_SCRIPTS_COMPARTILHADOS
python validar_reorganizacao.py
```

**O script valida:**
- ✅ Todas as bases SSOT existem
- ✅ Todos os projetos conseguem acessar as bases
- ✅ Estrutura de diretórios está íntegra
- ✅ Scripts compartilhados funcionais

---

### 📊 4. Gerar Inventário Atualizado

Para catalogar todos os arquivos do ecossistema:

```bash
cd C:\Users\a483650\Projetos\Controle_de_entregas
python inventario_projetos.py
```

**Output:**
- `INVENTARIO_PROJETOS_DATA_ANALYTICS.xlsx` atualizado
- Catálogo de 1.581+ arquivos
- Tamanhos, tipos, datas de modificação

---

### 🔎 5. Analisar Duplicações (Manutenção)

Para verificar se surgiram novas duplicações:

```bash
cd C:\Users\a483650\Projetos\Controle_de_entregas
python analisar_duplicacoes.py
```

**Output:**
- `RELATORIO_DUPLICACOES.xlsx` atualizado
- Análise MD5 de todos os arquivos
- Grupos de duplicatas identificados
- Recomendações de limpeza

---

## 🎯 PRINCÍPIOS DE USO

### 🔑 Regra de Ouro: SSOT (Single Source of Truth)
**SEMPRE** use as bases de `_DADOS_CENTRALIZADOS/`  
**NUNCA** copie bases manualmente para projetos  
**SEMPRE** execute `sincronizar_bases.py` após atualizar bases

### 🔄 Fluxo de Atualização
```
1. Nova base chega
2. Versão antiga → historico/ com data
3. Nova versão → renomear para _ATUAL
4. Executar sincronizar_bases.py
5. Validar com validar_reorganizacao.py
```

### 📁 Nomenclatura de Arquivos
- **ATUAL**: Versão oficial em produção (`hubspot_leads_ATUAL.csv`)
- **Histórico**: Versões antigas com data (`hubspot_leads_2026-02-10.csv`)
- **Outputs**: Sempre incluir data/timestamp nos outputs

### 🚫 O que NÃO fazer
❌ Copiar bases manualmente entre projetos  
❌ Editar arquivos em `_DADOS_CENTRALIZADOS/` diretamente  
❌ Deletar arquivos sem fazer backup em `historico/`  
❌ Criar novas pastas sem consultar a arquitetura  

---

## 📚 DOCUMENTAÇÃO DISPONÍVEL

### Para Uso Diário
- **GUIA_USO.md** (este arquivo) - Manual operacional
- **README.md** - Visão geral do ecossistema

### Para Entender a Arquitetura
- **PROJETO_DATA_SCIENCE.md** - Cronograma e desenvolvimento
- **ARQUITETURA_ECOSSISTEMA.md** - Princípios e design técnico

### Para Apresentações
- **RESUMO_EXECUTIVO.md** - Resumo executivo completo
- **RESUMO_EXECUTIVO_1PAGINA.md** - Resumo de 1 página
- **SCRIPT_APRESENTACAO.md** - Guia slide-a-slide
- **CHECKLIST_APRESENTACAO.md** - Preparação rápida

### Para Inventários
- **INDICE_ARQUIVOS.md** - Índice completo de arquivos
- **INVENTARIO_PROJETOS_DATA_ANALYTICS.xlsx** - Catálogo de 1.581 arquivos
- **CONTROLE_ENTREGAS_ECOSSISTEMA_DATA_SCIENCE.xlsx** - Planilha mestre

---

## ⚡ ATALHOS RÁPIDOS

### Atualização Express
```powershell
# 1. Atualizar base + sincronizar + validar
cd C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot
Move-Item hubspot_leads_ATUAL.csv historico/hubspot_leads_$(Get-Date -Format 'yyyy-MM-dd').csv
Copy-Item C:\Downloads\novo_file.csv hubspot_leads_ATUAL.csv
cd ..\..\SCRIPTS_COMPARTILHADOS
python sincronizar_bases.py
python validar_reorganizacao.py
```

### Executar Ambos Projetos de Produção
```powershell
# Helio ML
cd C:\Users\a483650\Projetos\01_Helio_ML_Producao\Scritps
python 1.ML_Lead_Scoring.py

# Pipeline Meta
cd ..\..\02_Pipeline_Midia_Paga\scripts
python pipeline_completo.py
```

---

## 🆘 TROUBLESHOOTING

### ❓ Script não encontra base de dados
**Causa:** Base não sincronizada ou caminho incorreto  
**Solução:**
```bash
python _SCRIPTS_COMPARTILHADOS/sincronizar_bases.py
python _SCRIPTS_COMPARTILHADOS/validar_reorganizacao.py
```

### ❓ Erro de encoding ao ler CSV
**Causa:** CSV com encoding diferente de UTF-8  
**Solução:** Adicione ao código:
```python
pd.read_csv(file, encoding='utf-8-sig', errors='replace')
```

### ❓ Arquivo muito grande trava script
**Causa:** Arquivo 29+ MB carregado de uma vez  
**Solução:** Use chunking:
```python
for chunk in pd.read_csv(file, chunksize=10000):
    # processar chunk
```

### ❓ Script de sincronização não funciona
**Causa:** Paths absolutos quebrados  
**Solução:** Verifique paths em `sincronizar_bases.py` - devem ser absolutos ou usar `pathlib.Path`

---

## 📊 MÉTRICAS DO SISTEMA

### Status Atual (Fevereiro 2026)
- ✅ **4 scripts** de automação 100% funcionais
- ✅ **2 projetos ML** em produção
- ✅ **5 bases oficiais** centralizadas (41 MB)
- ✅ **1.581 arquivos** catalogados
- ✅ **1.09 GB** de dados otimizados
- ✅ **Zero** duplicações críticas
- ✅ **12 documentos** profissionais

### Frequência de Uso
- **Sincronização:** Semanal ou quando novas bases chegam
- **Validação:** Após qualquer mudança estrutural
- **Lead Scoring:** Semanal ou sob demanda
- **Pipeline Meta:** Semanal ou início de mês
- **Inventário:** Mensal ou trimestral

---

## 🔮 PRÓXIMOS PASSOS

### Curto Prazo (Q1 2026)
- [ ] Dashboard interativo com Streamlit
- [ ] Migração de CSV → PostgreSQL
- [ ] API REST para acesso aos dados
- [ ] Testes automatizados (pytest)

### Médio Prazo (Q2-Q3 2026)
- [ ] CI/CD com GitHub Actions
- [ ] Monitoring de pipelines
- [ ] MLflow para modelos
- [ ] Documentação self-service

---

## 👤 CONTATO

**Responsável:** Data Science Team  
**Início:** Outubro 2025  
**Status:** ✅ Operacional desde Fevereiro 2026

Para dúvidas técnicas, consulte:
1. Este guia (uso diário)
2. `ARQUITETURA_ECOSSISTEMA.md` (design técnico)
3. `README.md` (visão geral)

---

*Sistema criado com foco em profissionalismo, escalabilidade e impacto de negócio.*
