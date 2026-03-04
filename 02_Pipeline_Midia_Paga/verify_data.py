import pandas as pd
from pathlib import Path

resumido_path = Path("outputs/visão resumida/hubspot_resumido_Dash.xlsx")

print("Verificando arquivo resumido...\n")

# Ler as abas
df_blend = pd.read_excel(resumido_path, sheet_name="Blend_Agregado_Dash")
df_granular = pd.read_excel(resumido_path, sheet_name="Visao_Granular_Final")
df_matriculas = pd.read_excel(resumido_path, sheet_name="Matricula_Data_Fechamento")

print("=" * 70)
print("ABA 1 - Blend_Agregado_Dash (Data de Criação)")
print("=" * 70)
df_blend['Ano'] = pd.to_datetime(df_blend['Data_Criacao']).dt.year
df_blend['Mes'] = pd.to_datetime(df_blend['Data_Criacao']).dt.month

jan_2026_blend = df_blend[(df_blend['Ano'] == 2026) & (df_blend['Mes'] == 1)]
print(f"Janeiro 2026: {jan_2026_blend['Volume_Total_Negocios'].sum():.0f} leads")

print("\n" + "=" * 70)
print("ABA 2 - Visao_Granular_Final (Data de Criação)")
print("=" * 70)  
df_granular['Ano'] = pd.to_datetime(df_granular['Data']).dt.year
df_granular['Mes'] = pd.to_datetime(df_granular['Data']).dt.month

jan_2026_granular = df_granular[(df_granular['Ano'] == 2026) & (df_granular['Mes'] == 1)]
print(f"Janeiro 2026: {len(jan_2026_granular)} leads")

print("\n" + "=" * 70)
print("ABA 3 - Matricula_Data_Fechamento (Data de Fechamento)")
print("=" * 70)
df_matriculas['Ano'] = pd.to_datetime(df_matriculas['Data_Fechamento'], errors='coerce').dt.year
df_matriculas['Mes'] = pd.to_datetime(df_matriculas['Data_Fechamento'], errors='coerce').dt.month

jan_2026_mat = df_matriculas[(df_matriculas['Ano'] == 2026) & (df_matriculas['Mes'] == 1)]
print(f"Janeiro 2026: {len(jan_2026_mat)} matrículas (por data de fechamento)")
print("=" * 70)
