import pandas as pd
from pathlib import Path

excel_file = Path(r"c:\Users\a483650\Projetos\06_Analise_Funil_RedBalloon\outputs\funil_redballoon_20260211_174832.xlsx")

# Verifica aba comparativa COM TODAS AS COLUNAS
print("=" * 80)
print("ABA 7b_Comparativo_26vs25 - COMPLETA")
print("=" * 80)
df_comp = pd.read_excel(excel_file, sheet_name="7b_Comparativo_26vs25")
print(f"Colunas: {df_comp.columns.tolist()}")
print("\nDados completos:")
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
print(df_comp)

# Filtra Meta Ads
print("\n" + "=" * 80)
print("META ADS - CICLOS 25.1 e 26.1")
print("=" * 80)
meta_ads = df_comp[df_comp['Canal'] == 'Meta Ads']
print(meta_ads)
