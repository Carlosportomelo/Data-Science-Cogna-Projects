import pandas as pd
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
RESUMIDO = BASE / 'outputs' / 'visão resumida' / 'hubspot_resumido_Dash.xlsx'
OUT_DIR = BASE / 'outputs' / 'visão resumida' / 'heuristics'
OUT_DIR.mkdir(parents=True, exist_ok=True)

print('Carregando resumido:', RESUMIDO)
df = pd.read_excel(RESUMIDO, sheet_name='Visao_Granular_Final')

# Normalizações
for col in ['Lead_Key', 'Origem_Principal']:
    if col in df.columns:
        df[col] = df[col].astype(str)

if 'Matriculas' in df.columns:
    df['Matriculas'] = pd.to_numeric(df['Matriculas'], errors='coerce').fillna(0).astype(int)
else:
    df['Matriculas'] = 0

# Parse dates
if 'Data_Fechamento' in df.columns:
    df['Data_Fechamento'] = pd.to_datetime(df['Data_Fechamento'], errors='coerce')
else:
    df['Data_Fechamento'] = pd.NaT

if 'Data' in df.columns:
    df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
else:
    df['Data'] = pd.NaT

# Helper
def jan26_mask_on(col):
    return (col.dt.year == 2026) & (col.dt.month == 1)

# Heurística A: Data_Fechamento estrito
a = df[(df['Matriculas'] == 1) & (df['Data_Fechamento'].notna()) & jan26_mask_on(df['Data_Fechamento'])].copy()
# Heurística B: Matriculas==1 usando Data (criação) - apenas para comparação
b = df[(df['Matriculas'] == 1) & (df['Data'].notna()) & jan26_mask_on(df['Data'])].copy()
# Heurística C: Fallback (preencher Data_Fechamento com Data para status MATR*)
mask_status_matr = df['Status_Principal'].astype(str).str.upper().str.contains('MATR', na=False)
df_c = df.copy()
fill_idx = mask_status_matr & df_c['Data_Fechamento'].isna()
df_c.loc[fill_idx, 'Data_Fechamento'] = df_c.loc[fill_idx, 'Data']
c = df_c[(df_c['Matriculas'] == 1) & (df_c['Data_Fechamento'].notna()) & jan26_mask_on(df_c['Data_Fechamento'])].copy()
# Heurística D: Matriculas==1 e Midia_Paga>0
if 'Midia_Paga' in df.columns:
    d = df[(df['Matriculas'] == 1) & (pd.to_numeric(df['Midia_Paga'], errors='coerce') > 0) & (df['Data_Fechamento'].notna()) & jan26_mask_on(df['Data_Fechamento'])].copy()
else:
    d = df.iloc[0:0].copy()
# Heurística E: Canais pagos apenas (Social Pago + Pesquisa Paga) com Data_Fechamento
paid_ch = ['Social Pago', 'Pesquisa Paga']
e = df[(df['Matriculas'] == 1) & (df['Origem_Principal'].isin(paid_ch)) & (df['Data_Fechamento'].notna()) & jan26_mask_on(df['Data_Fechamento'])].copy()

results = {
    'data_fechamento_strict': a,
    'matriculas_on_data_criacao': b,
    'fallback_status_fill': c,
    'matriculas_with_midia': d,
    'paid_channels_data_fech': e
}

summary_lines = []
for name, subset in results.items():
    unique_leads = subset['Lead_Key'].nunique() if 'Lead_Key' in subset.columns else subset.shape[0]
    rows = subset.shape[0]
    summary_lines.append((name, rows, unique_leads))
    out_csv = OUT_DIR / f'{name}_jan2026.csv'
    subset.to_csv(out_csv, index=False, encoding='utf-8-sig')
    print(f'Exported {name} -> {out_csv} (rows: {rows}, unique_leads: {unique_leads})')

print('\nResumo comparativo:')
for line in summary_lines:
    print(f' - {line[0]}: rows={line[1]}, unique_leads={line[2]}')

# Also print difference between fallback and strict
strict_keys = set(a['Lead_Key'].unique()) if not a.empty else set()
fallback_keys = set(c['Lead_Key'].unique()) if not c.empty else set()
missing_in_strict = fallback_keys - strict_keys
if missing_in_strict:
    print(f"\nLeads presentes com fallback mas ausentes no estrito (count={len(missing_in_strict)}). Exemplos:")
    for i, k in enumerate(list(missing_in_strict)[:50]):
        print(k)
    with open(OUT_DIR / 'missing_in_strict.txt', 'w', encoding='utf-8') as fh:
        for k in missing_in_strict:
            fh.write(k + '\n')
    print(f"Lista exportada: {OUT_DIR / 'missing_in_strict.txt'}")
else:
    print('\nNenhum lead adicional encontrado pelo fallback.')
