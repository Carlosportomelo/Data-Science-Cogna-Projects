import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# Carregar os dados
file_path = r'c:\Users\a483650\Projetos\03_Analises_Operacionais\ICP\DATA\Escolas_Novas_2025.xlsx'

# Primeiro, vamos ver quais sheets existem
excel_file = pd.ExcelFile(file_path)
print("=" * 80)
print("SHEETS DISPONÍVEIS NO ARQUIVO:")
print("=" * 80)
for sheet in excel_file.sheet_names:
    print(f"- {sheet}")
print("\n")

# Carregar todas as sheets para exploração
sheets_data = {}
for sheet in excel_file.sheet_names:
    sheets_data[sheet] = pd.read_excel(file_path, sheet_name=sheet)
    print(f"=" * 80)
    print(f"SHEET: {sheet}")
    print(f"=" * 80)
    print(f"Dimensões: {sheets_data[sheet].shape}")
    print(f"\nColunas disponíveis:")
    for col in sheets_data[sheet].columns:
        print(f"  - {col}")
    print(f"\nPrimeiras linhas:")
    print(sheets_data[sheet].head())
    print("\n")
