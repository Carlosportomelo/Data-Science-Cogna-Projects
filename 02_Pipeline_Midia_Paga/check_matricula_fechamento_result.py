import pandas as pd
from pathlib import Path

out_dir = Path('outputs')
files = sorted(out_dir.glob('meta_googleads_blend_*.xlsx'), key=lambda p: p.stat().st_mtime, reverse=True)
if not files:
    print('Nenhum arquivo gerado encontrado.')
    raise SystemExit(1)

latest = files[0]
print(f'Arquivo: {latest}')

sheet = 'Matricula_Data_Fechamento'
df = pd.read_excel(latest, sheet_name=sheet)

if 'Data_Fechamento' not in df.columns:
    print('Coluna Data_Fechamento não encontrada na aba.')
    raise SystemExit(1)

df['Data_Fechamento'] = pd.to_datetime(df['Data_Fechamento'], errors='coerce')
dec = df[(df['Data_Fechamento'] >= '2025-12-01') & (df['Data_Fechamento'] < '2026-01-01')]

print(f'Dezembro/2025 ({sheet}): {len(dec)} matrículas')
if 'Origem_Principal' in dec.columns:
    print('Por canal:')
    print(dec['Origem_Principal'].value_counts().to_string())
