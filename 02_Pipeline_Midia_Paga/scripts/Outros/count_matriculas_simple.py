import pandas as pd
P='outputs/visão resumida/hubspot_resumido_Dash.xlsx'
try:
    df=pd.read_excel(P, sheet_name='Visao_Granular_Final', usecols=['Matriculas'])
except Exception as e:
    print('Erro ao ler coluna Matriculas:', e)
    # tentar ler todas e filtrar
    df=pd.read_excel(P, sheet_name='Visao_Granular_Final')
    if 'Matriculas' not in df.columns:
        print('Coluna Matriculas não encontrada nas colunas:', list(df.columns))
        raise SystemExit

s = pd.to_numeric(df['Matriculas'], errors='coerce').fillna(0).sum()
rows_with = (pd.to_numeric(df['Matriculas'], errors='coerce').fillna(0) > 0).sum()
print(f'Soma Matriculas: {int(s)}')
print(f'Linhas com Matriculas>0: {int(rows_with)}')
