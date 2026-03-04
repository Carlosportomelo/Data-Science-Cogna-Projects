import pandas as pd
from pathlib import Path

excel_file = Path(r"c:\Users\a483650\Projetos\06_Analise_Funil_RedBalloon\outputs\funil_redballoon_20260211_174832.xlsx")

# Lista abas
xl = pd.ExcelFile(excel_file)
print("ABAS DISPONÍVEIS:")
print("=" * 80)
for i, aba in enumerate(xl.sheet_names, 1):
    print(f"{i}. {aba}")

# Verifica aba 7_Matriz_Ciclo_Fonte
print("\n" + "=" * 80)
print("ABA 7_Matriz_Ciclo_Fonte")
print("=" * 80)
df_matriz = pd.read_excel(excel_file, sheet_name="7_Matriz_Ciclo_Fonte")
print(f"Colunas: {df_matriz.columns.tolist()}")
print(f"\nShape: {df_matriz.shape}")
print("\nPrimeiras linhas:")
print(df_matriz.head(20))

# Tenta encontrar dados do ciclo 25.1 com Meta Ads
print("\n" + "=" * 80)
print("BUSCANDO CICLO 25.1 COM META ADS")
print("=" * 80)
# A aba pode ser uma matriz pivotada, então vamos ver a estrutura
if 'Meta Ads' in df_matriz.columns:
    print("\nMeta Ads é uma COLUNA. Buscando linha 25.1:")
    if '25.1' in df_matriz.index or '25.1' in df_matriz.iloc[:, 0].values:
        linha_251 = df_matriz[df_matriz.iloc[:, 0] == '25.1']
        print(linha_251)
else:
    print("\nMeta Ads NÃO é coluna. Estrutura:")
    print(df_matriz.head(30))
