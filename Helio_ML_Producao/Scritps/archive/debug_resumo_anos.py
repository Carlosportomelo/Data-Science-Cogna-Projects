import pandas as pd
import os
from datetime import datetime

CAMINHO_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CAMINHO_ENTRADA = os.path.join(CAMINHO_BASE, 'Data', 'hubspot_leads.csv')

try:
    df = pd.read_csv(CAMINHO_ENTRADA, encoding='utf-8')
except:
    df = pd.read_csv(CAMINHO_ENTRADA, encoding='latin1')

# Replicar o parsing robusto do script principal
raw_dates = df['Data de criação'].astype(str)
parsed_default = pd.to_datetime(raw_dates, errors='coerce')
parsed_dayfirst = pd.to_datetime(raw_dates, dayfirst=True, errors='coerce')
df['Data de criação'] = parsed_default.fillna(parsed_dayfirst)
df['Ano'] = df['Data de criação'].dt.year

print('Total leads:', len(df))
print('Anos detectados (sorted):', sorted(df['Ano'].dropna().unique()))
print('\nContagem por ano:\n', df['Ano'].value_counts(dropna=False).sort_index())

resumo = df.groupby('Ano').agg(Total_Leads=('Ano', 'count'), Matriculas=('Etapa do negócio', lambda s: s.str.contains('MATRÍCULA', case=False, na=False).sum())).reset_index()
resumo['Taxa_Conversao'] = resumo['Matriculas'] / resumo['Total_Leads']

print('\nResumo gerado:\n', resumo.sort_values('Ano'))

# Mostrar exemplos de datas que não foram parseadas
na_dates = df[df['Data de criação'].isna()]['Data de criação'].index[:10]
if len(na_dates) > 0:
    print('\nExemplos de linhas com Data de criação inválida (índices):', list(na_dates))
