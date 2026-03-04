import pandas as pd

P = r"outputs/visão resumida/meta_resumido_Dash.xlsx"
xls = pd.ExcelFile(P)
found = False
for sheet in xls.sheet_names:
    df = pd.read_excel(P, sheet_name=sheet)
    date_col = None
    value_col = None
    for c in df.columns:
        cl = str(c).lower()
        if date_col is None and 'data' in cl:
            date_col = c
        if value_col is None and any(k in cl for k in ['invest', 'valor', 'custo', 'cost']):
            value_col = c
    if date_col and value_col:
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df[value_col] = pd.to_numeric(df[value_col].astype(str).str.replace(',', ''), errors='coerce')
        s = df.loc[(df[date_col].dt.year == 2026) & (df[date_col].dt.month == 1), value_col].sum()
        print(f"{sheet} | {date_col} | {value_col} | R$ {s:,.2f}")
        found = True
    else:
        print(f"{sheet} | sem colunas Data/Investimento detectadas")

if not found:
    print("Nenhum valor encontrado para janeiro/2026 neste arquivo.")
