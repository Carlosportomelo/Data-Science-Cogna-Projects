import pandas as pd
from pathlib import Path
BASE = Path(__file__).resolve().parent.parent
CSV = BASE / 'outputs' / 'visão resumida' / 'heuristics' / 'matriculas_on_data_criacao_jan2026.csv'
OUT = BASE / 'outputs' / 'visão resumida' / 'heuristics' / 'matriculas_on_data_paid_subset.csv'
print('Lendo:', CSV)
df = pd.read_csv(CSV, encoding='utf-8-sig')
print('\nContagem por Origem_Principal:')
print(df['Origem_Principal'].value_counts(dropna=False))
paid = df[df['Origem_Principal'].isin(['Social Pago','Pesquisa Paga'])].copy()
paid.to_csv(OUT, index=False, encoding='utf-8-sig')
print(f'Exportado subset pago: {OUT} (rows: {len(paid)})')
