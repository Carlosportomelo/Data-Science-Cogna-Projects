import pandas as pd
P='outputs/visão resumida/hubspot_resumido_Dash.xlsx'
try:
    df=pd.read_excel(P, sheet_name='Visao_Granular_Final')
except Exception:
    df=pd.read_excel(P)
cols=[c for c in df.columns]
mat_cols=[c for c in cols if 'matricul' in str(c).lower()]
if not mat_cols:
    print('Nenhuma coluna com "matricul" encontrada nas colunas:', cols)
else:
    total_by_col={}
    for c in mat_cols:
        s=pd.to_numeric(df[c], errors='coerce').fillna(0).sum()
        total_by_col[c]=s
    df_flags = df[mat_cols].apply(lambda x: pd.to_numeric(x, errors='coerce').fillna(0))
    rows_with_mat = (df_flags.sum(axis=1) >= 1).sum()
    primary_sum = None
    if 'Matriculas' in df.columns:
        primary_sum = pd.to_numeric(df['Matriculas'], errors='coerce').fillna(0).sum()
    print('Colunas de matriculas encontradas:', mat_cols)
    for k,v in total_by_col.items():
        print(f"Soma {k}: {int(v)}")
    if primary_sum is not None:
        print(f"Soma coluna 'Matriculas': {int(primary_sum)}")
    print(f"Linhas com pelo menos 1 matricula marcada: {int(rows_with_mat)}")
