import pandas as pd
from pathlib import Path

csv = r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_ATUAL.csv"
out = Path(r"06_Analise_Funil_RedBalloon/outputs/Analise_Ciclo_26_1_POR_UNIDADE.xlsx")

df = pd.read_csv(csv, low_memory=False)
# parse dates
for c in ['Data de criação','Data da última atividade','Data de fechamento']:
    if c in df.columns:
        df[c] = pd.to_datetime(df[c], errors='coerce')

# definir_ciclo
def definir_ciclo(date):
    if pd.isna(date):
        return None
    ano = date.year
    mes = date.month
    if mes >= 10:
        return f"{str(ano+1)[-2:]}.1"
    elif mes <= 3:
        return f"{str(ano)[-2:]}.1"
    else:
        return None

if 'Data de criação' not in df.columns:
    print('Coluna Data de criação não encontrada')
    raise SystemExit(1)

df['ciclo'] = df['Data de criação'].apply(definir_ciclo)

df['converteu'] = (df['Etapa do negócio'].astype(str).str.upper().str.contains('MATRÍCULA CONCLUÍDA', na=False)) | (df['Data de fechamento'].notna())

# filtro ciclo
cdf = df[df['ciclo']=='26.1'].copy()
# totals
total_leads = len(cdf)
# leads valid
cdf_valid = cdf[(cdf['Data da última atividade'].notna()) & (cdf['Data da última atividade'] >= cdf['Data de criação'])].copy()
leads_valid = len(cdf_valid)
# matriculas
mats = cdf[cdf['converteu']==True].copy()
total_matriculas = len(mats)
# matriculas valid
mats_valid = mats[(mats['Data da última atividade'].notna()) & (mats['Data da última atividade'] >= mats['Data de criação'])].copy()
matriculas_valid = len(mats_valid)

print('Recalculo direto da base:')
print(f' Total Leads (ciclo 26.1): {total_leads}')
print(f' Leads válidos (última atividade >= criação): {leads_valid}')
print(f' Total Matrículas (ciclo 26.1): {total_matriculas}')
print(f' Matrículas válidas (última atividade >= criação): {matriculas_valid}')

# agora ler arquivo por unidade e somar Totais por unidade
if out.exists():
    df_units = pd.read_excel(out, sheet_name='Unidades (Leads)')
    # identificar linhas de total por canal where Canal contains 'TOTAL '
    if 'Canal' in df_units.columns and 'Total Leads' in df_units.columns:
        unit_totals = df_units[df_units['Canal'].astype(str).str.contains('TOTAL')]
        s_unit_totals = unit_totals['Total Leads'].sum()
        s_all_rows = df_units['Total Leads'].sum()
        print('\nComparação com arquivo Analise_Ciclo_26_1_POR_UNIDADE.xlsx (Unidades (Leads)):')
        print(f' Soma de Total Leads (todas linhas): {s_all_rows}')
        print(f' Soma de Total Leads (linhas de TOTAL por unidade): {s_unit_totals}')
        print(f' Diferença (recalculo base - soma totais unidades): {total_leads - s_unit_totals}')
    else:
        print('Colunas esperadas não encontradas na aba Unidades (Leads)')
else:
    print('Arquivo POR_UNIDADE.xlsx não encontrado:', out)
