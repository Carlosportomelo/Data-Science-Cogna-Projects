# 01_Helio_ML_Producao

Projeto de machine learning para scoring de leads, validacao de conversao e geracao de relatorios executivos.

## Objetivo
- Priorizar leads com maior probabilidade de conversao.
- Suportar operacao com analises tecnicas e gerenciais.
- Padronizar execucao do pipeline de scoring.

## Estrutura
- `Scritps/`: scripts principais do pipeline e utilitarios.
- `Docs/`: metodologia e notas de implementacao.
- `models/`: artefatos de modelo.
- `Data/` e `Outputs/`: dados e saidas (ignorado no Git quando sensivel).

## Scripts-chave
- `Scritps/0.Master_Pipeline.py`: orquestracao principal.
- `Scritps/1.ML_Lead_Scoring.py`: scoring dos leads.
- `Scritps/10.Validacao_Conversao_Helio.py`: validacao de conversao.
- `Scritps/6.Relatorio_Executivo.py`: consolidacao executiva.

## Como executar
```bash
cd 01_Helio_ML_Producao/Scritps
python 0.Master_Pipeline.py
```

## Periodo
Outubro/2025 a Marco/2026.
