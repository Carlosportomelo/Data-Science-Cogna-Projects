import pandas as pd
import os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PATH = os.path.join(BASE, 'Data', 'hubspot_negocios_perdidos.csv')

try:
    df = pd.read_csv(PATH, encoding='utf-8')
except Exception:
    df = pd.read_csv(PATH, encoding='latin1')

col = 'Etapa do negócio'
if col not in df.columns:
    print(f"Coluna '{col}' não encontrada. Cabeçalho: {df.columns.tolist()}")
else:
    vals = df[col].astype(str).str.strip().unique()
    print(f"Unique Etapa do negócio (count={len(vals)}):")
    for v in sorted(vals, key=lambda x: (str(x).lower())):
        print('-', v)
