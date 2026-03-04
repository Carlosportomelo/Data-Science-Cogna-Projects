# Análise Comparativa - Leads Offline e Tráfego Direto

## 📋 Descrição

Este projeto analisa o crescimento de leads provenientes de canais **Offline** e **Tráfego Direto** comparando os **relatórios consolidados** já gerados em datas diferentes.

**IMPORTANTE**: Este script lê os relatórios consolidados existentes na pasta `analise_consolidada_leads_matriculas` e extrai apenas os dados de canais Offline e Tráfego Direto para análise comparativa.

## 🎯 Objetivo

Comparar os dados de leads offline e tráfego direto entre os relatórios gerados nas datas:
- **30/01/2026** (30 de janeiro)
- **06/02/2026** (6 de fevereiro) 
- **09/02/2026** (9 de fevereiro)

## 📁 Estrutura do Projeto

```
crescimento_offline_direto/
├── data/               # Dados de entrada (não utilizada - lê de analise_consolidada_leads_matriculas)
├── outputs/            # Relatórios gerados
│   ├── comparativo_offline_direto_YYYYMMDD_HHMMSS.xlsx
│   └── relatorio_comparativo_YYYYMMDD_HHMMSS.txt
├── scripts/            # Scripts de análise
│   └── analise_crescimento_offline_direto.py
├── README.md           # Este arquivo
└── requirements.txt    # Dependências
```

## 🔍 O que é analisado

### Fonte de Dados
Lê os seguintes relatórios consolidados:
- `analise_consolidada_leads_matriculas_2026-01-30.xlsx`
- `analise_consolidada_leads_matriculas_2026-02-06.xlsx`
- `analise_consolidada_leads_matriculas_2026-02-09.xlsx`

### Canais Analisados
- **Offline**: "Fontes off-line"
- **Tráfego Direto**: "Tráfego direto"

### Métricas Principais
1. **Comparação por ciclo** (23.1, 24.1, 25.1, 26.1)
2. **Total Geral** (acumulado de todos os ciclos)
3. **Crescimento absoluto** entre datas (ex: +96 leads)
4. **Crescimento percentual** entre datas (ex: +2.0%)
5. **Análise específica do Ciclo 26.1** (ciclo mais recente)

## 📊 Relatórios Gerados

### 1. Arquivo Excel (`.xlsx`)
Contém 4 abas organizadas:

#### Aba 1: Resumo Executivo
- Dados do Ciclo 26.1 (mais recente) para as 3 datas
- Total Geral (acumulado) para as 3 datas
- Visão rápida e consolidada

#### Aba 2: Comparativo por Ciclo
- Todos os ciclos (23.1, 24.1, 25.1, 26.1)
- Total Geral
- Valores de Offline, Tráfego Direto e Total
- Para cada data de relatório

#### Aba 3: Crescimento Detalhado
- Crescimento entre cada par de datas
- Para todos os ciclos e métricas
- Valores absolutos e percentuais

#### Aba 4: Crescimento Ciclo 26.1
- Foco específico no ciclo mais recente
- Crescimento de Offline, Tráfego Direto e Total
- Formatação destacada para fácil visualização

### 2. Relatório Texto (`.txt`)
- Resumo executivo em formato texto
- Ciclo 26.1 e Total Geral
- Fácil visualização rápida

## 🚀 Como Usar

### Pré-requisitos
```bash
# Bibliotecas necessárias
pip install pandas numpy openpyxl
```

Ou use o arquivo requirements.txt:
```bash
pip install -r requirements.txt
```

### Execução

1. **Via PowerShell:**
```powershell
cd c:\Users\a483650\Projetos\03_Analises_Operacionais\crescimento_offline_direto\scripts
python analise_crescimento_offline_direto.py
```

2. **Usando o ambiente virtual do projeto:**
```powershell
cd c:\Users\a483650\Projetos\03_Analises_Operacionais\crescimento_offline_direto\scripts
C:/Users/a483650/Projetos/03_Analises_Operacionais/.venv/Scripts/python.exe analise_crescimento_offline_direto.py
```

### Saída Esperada
O script irá:
1. Carregar os 3 relatórios consolidados
2. Extrair linhas de "Fontes off-line" e "Tráfego direto"
3. Comparar valores entre as datas
4. Calcular crescimento absoluto e percentual
5. Gerar relatórios Excel e texto
6. Exibir resumo no console

## 📈 Exemplo de Análise

```
CICLO 26.1 (Mais Recente):
Data          Offline    Tráfego Direto    Total
30/01/2026     4,453          809          5,262
06/02/2026     4,821          852          5,673
09/02/2026     4,917          861          5,778

Crescimento 06/02 → 09/02:
- Offline: +96 (+2.0%)
- Tráfego Direto: +9 (+1.1%)
- Total: +105 (+1.9%)
```

## 🔧 Configuração

### Modificar Relatórios Analisados
Edite o dicionário `RELATORIOS` no script:
```python
RELATORIOS = {
    '30/01/2026': 'analise_consolidada_leads_matriculas_2026-01-30.xlsx',
    '06/02/2026': 'analise_consolidada_leads_matriculas_2026-02-06.xlsx',
    '09/02/2026': 'analise_consolidada_leads_matriculas_2026-02-09.xlsx'
}
```

### Modificar Canais
Edite as variáveis:
```python
CANAL_OFFLINE = 'Fontes off-line'
CANAL_DIRETO = 'Tráfego direto'
```

## 📝 Observações Importantes

1. **Fonte de Dados**: O script lê os relatórios consolidados da pasta `eficiencia_canal/outputs/analise_consolidada_leads_matriculas/`
2. **Aba Lida**: Sempre lê a aba "Leads (Canais)" de cada relatório
3. **Nomes Exatos**: Os nomes dos canais devem ser exatamente como aparecem nos relatórios
4. **Timestamp**: Arquivos gerados incluem timestamp único para evitar sobrescrita

## 🐛 Troubleshooting

### Erro: "Arquivo não encontrado"
- Verifique se os arquivos dos relatórios existem na pasta `analise_consolidada_leads_matriculas`
- Confirme os nomes exatos dos arquivos no dicionário `RELATORIOS`

### Erro: "Canal não encontrado"
- Verifique se os nomes dos canais (`CANAL_OFFLINE` e `CANAL_DIRETO`) estão exatamente como nos relatórios
- Cuidado com maiúsculas/minúsculas e hifens

### Erro de importação
```bash
# Instale as dependências
pip install -r requirements.txt
```

## 📧 Suporte

Para dúvidas ou problemas:
1. Verifique se os relatórios consolidados existem
2. Confirme os nomes dos arquivos e canais
3. Revise os logs no console durante a execução

## 📅 Histórico de Versões

- **v2.0** (10/02/2026): Versão corrigida que lê relatórios consolidados existentes
- **v1.0** (10/02/2026): Versão inicial (descontinuada)

---

**Projeto**: Análises Operacionais Red Balloon  
**Autor**: Equipe de Análise Operacional  
**Última atualização**: 10/02/2026
