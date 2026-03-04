import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# Carregar os dados
file_path = r'c:\Users\a483650\Projetos\03_Analises_Operacionais\ICP\DATA\Escolas_Novas_2025.xlsx'
df = pd.read_excel(file_path, sheet_name='Escolas')

print("=" * 80)
print("EXPLORAÇÃO DETALHADA DOS DADOS")
print("=" * 80)
print(f"\nTotal de escolas: {len(df)}")
print(f"\nColunas: {list(df.columns)}")
print("\n" + "=" * 80)
print("AMOSTRA DOS DADOS (primeiras 10 linhas):")
print("=" * 80)
print(df.head(10).to_string())
print("\n" + "=" * 80)
print("INFORMAÇÕES SOBRE TIPOS DE DADOS E VALORES ÚNICOS:")
print("=" * 80)
for col in df.columns:
    print(f"\n{col}:")
    print(f"  Tipo: {df[col].dtype}")
    print(f"  Valores únicos: {df[col].nunique()}")
    print(f"  Valores nulos: {df[col].isna().sum()}")
    if df[col].nunique() < 20 and col not in ['Nome da Escola', 'Observação']:
        print(f"  Valores únicos: {df[col].unique()[:10]}")
