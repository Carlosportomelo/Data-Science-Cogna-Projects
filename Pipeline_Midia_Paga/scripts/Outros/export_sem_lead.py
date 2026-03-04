import pandas as pd
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
P = BASE / "outputs" / "visão resumida" / "hubspot_resumido_Dash.xlsx"
OUT_DIR = BASE / "outputs" / "visão resumida"
OUT_DIR.mkdir(parents=True, exist_ok=True)

try:
    df = pd.read_excel(P, sheet_name='Visao_Granular_Final')
except Exception:
    # fallback: read first sheet
    df = pd.read_excel(P)

# normalizar nome de colunas possíveis
cols = [c for c in df.columns]
if 'Midia_Paga' not in df.columns:
    for c in cols:
        cl = str(c).lower()
        if 'midia' in cl or 'invest' in cl:
            df = df.rename(columns={c: 'Midia_Paga'})
            break

# garantir numérico
if 'Midia_Paga' in df.columns:
    df['Midia_Paga'] = pd.to_numeric(df['Midia_Paga'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
else:
    raise SystemExit('Coluna de investimento não encontrada (procure por Midia/Invest).')

# identificar linhas sem lead
unit_col = None
for c in df.columns:
    if str(c).lower() in ['unidade', 'unidade_desejada']:
        unit_col = c
        break

lead_key_col = None
for c in df.columns:
    if str(c).lower() in ['lead_key', 'lead key', 'leadkey']:
        lead_key_col = c
        break

mask = False
if unit_col:
    mask = df[unit_col].astype(str) == 'Sem Lead'
if lead_key_col:
    mask = mask | df[lead_key_col].astype(str).str.startswith('SEM_LEAD_')

# fallback: se nenhuma das heurísticas, procurar por 'Sem Lead' em todas as colunas
if not mask.any():
    mask = df.apply(lambda r: r.astype(str).str.contains('Sem Lead', case=False, na=False).any(), axis=1)

df_sem = df[mask].copy()

# salvar
csv_out = OUT_DIR / 'sem_lead_linhas.csv'
excel_out = OUT_DIR / 'sem_lead_linhas.xlsx'

df_sem.to_csv(csv_out, index=False, encoding='utf-8-sig')
df_sem.to_excel(excel_out, index=False)

print(f"Exportado {len(df_sem)} linhas 'Sem Lead' para:")
print(f" - CSV: {csv_out}")
print(f" - Excel: {excel_out}")
