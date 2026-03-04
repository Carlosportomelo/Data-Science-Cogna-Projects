import os
import glob
import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, 'Outputs')

files = glob.glob(os.path.join(OUT, 'Validacao_Fontes_Discriminadas_*.xlsx'))
if not files:
    print('Nenhum arquivo de validação encontrado em Outputs/')
    raise SystemExit(1)

# pegar o mais recente
latest = max(files, key=os.path.getmtime)
print('Usando arquivo:', latest)

df = pd.read_excel(latest, sheet_name='1_Resumo_Geral')
print('\nAba `1_Resumo_Geral` — primeiras linhas:')
print(df.head(10))

print('\nAnos presentes na aba:', sorted(df['Ano'].dropna().unique()))
