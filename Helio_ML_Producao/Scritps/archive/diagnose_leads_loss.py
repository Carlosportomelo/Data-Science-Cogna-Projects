import os
import pandas as pd

CAMINHO_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CAMINHO_ENTRADA = os.path.join(CAMINHO_BASE, 'Data', 'hubspot_leads.csv')

def count_file_lines(path):
    cnt = 0
    with open(path, 'r', encoding='utf-8', errors='replace') as f:
        for _ in f:
            cnt += 1
    return cnt

total_lines = count_file_lines(CAMINHO_ENTRADA)
print('Linhas no arquivo (inclui header):', total_lines)

try:
    df = pd.read_csv(CAMINHO_ENTRADA, encoding='utf-8')
except Exception:
    df = pd.read_csv(CAMINHO_ENTRADA, encoding='latin1')

print('Linhas lidas pelo pandas (DataFrame length):', len(df))

# Detect original nulls in the raw column before parsing
orig_nulls = df['Data de criação'].isna().sum()
print('Valores nulos originais em `Data de criação` antes do parse:', orig_nulls)

# Replicar parsing robusto
raw_dates = df['Data de criação'].astype(str)
parsed_default = pd.to_datetime(raw_dates, errors='coerce')
parsed_dayfirst = pd.to_datetime(raw_dates, dayfirst=True, errors='coerce')
final_dates = parsed_default.fillna(parsed_dayfirst)

parsed_count = final_dates.notna().sum()
print('Linhas com data parseada com sucesso:', parsed_count)
print('Linhas com data NÃO parseada (NaT):', final_dates.isna().sum())

diff_file_pandas = total_lines - len(df)
print('Diferença entre arquivo e DataFrame lido (linhas):', diff_file_pandas)

if final_dates.isna().sum() > 0:
    print('\nExemplos de linhas onde a data NÃO foi parseada:')
    bad_idx = final_dates[final_dates.isna()].index[:10]
    for i in bad_idx:
        raw = raw_dates.iat[i]
        # print a few key columns to help inspect
        row = df.iloc[i]
        etapa = str(row.get('Etapa do negócio'))[:40]
        fonte = str(row.get('Fonte original do tráfego'))
        print('Índice {}: raw_date="{}" | Etapa do negócio="{}" | Fonte="{}"'.format(i, raw, etapa, fonte))

print('\nResumo:')
print(f'  - Total linhas no CSV (incl header): {total_lines}')
print(f'  - Linhas lidas pelo pandas: {len(df)}')
print(f'  - Linhas com Ano detectado (parsed dates): {parsed_count}')
print(f'  - Perda potencial de linhas com Ano (não contadas no resumo): {final_dates.isna().sum()}')
