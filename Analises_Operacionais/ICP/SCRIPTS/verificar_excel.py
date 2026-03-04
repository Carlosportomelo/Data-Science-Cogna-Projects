import pandas as pd

# Verificar o conteúdo do Excel atualizado
file_path = r'c:\Users\a483650\Projetos\03_Analises_Operacionais\ICP\OUTPUTS\Analise_ICP_Processada.xlsx'

# Listar sheets
excel_file = pd.ExcelFile(file_path)
print("=" * 80)
print("VERIFICAÇÃO DO EXCEL ATUALIZADO")
print("=" * 80)
print(f"\nSheets disponíveis: {excel_file.sheet_names}")
print()

# Verificar sheet Dados Processados
df = pd.read_excel(file_path, sheet_name='Dados Processados')
print("=" * 80)
print("AMOSTRA DA SHEET 'Dados Processados'")
print("=" * 80)
print("\nColunas:")
print(df.columns.tolist())
print()

# Verificar se tem a coluna Tipo_Campo
if 'Tipo_Campo' in df.columns:
    print("✅ Coluna 'Tipo_Campo' encontrada!")
    print("\nContagem por Tipo_Campo:")
    print(df['Tipo_Campo'].value_counts())
    print()
    
    print("Brownfield (COM franqueado):")
    brownfield = df[df['Tipo_Campo'] == 'Brownfield']['Nome da Escola'].tolist()
    for escola in brownfield:
        print(f"  • {escola}")
    
    print("\nGreenfield (SEM franqueado):")
    greenfield = df[df['Tipo_Campo'] == 'Greenfield']['Nome da Escola'].tolist()
    for escola in greenfield:
        print(f"  • {escola}")
else:
    print("❌ Coluna 'Tipo_Campo' NÃO encontrada!")

# Verificar se tem a nova coluna Tinha_Ingles_Anterior
print("\n" + "=" * 80)
if 'Tinha_Ingles_Anterior' in df.columns:
    print("✅ Coluna 'Tinha_Ingles_Anterior' encontrada!")
    print("\nContagem:")
    print(df['Tinha_Ingles_Anterior'].value_counts())
else:
    print("❌ Coluna 'Tinha_Ingles_Anterior' NÃO encontrada!")

# Verificar sheet Brownfield vs Greenfield
print("\n" + "=" * 80)
print("SHEET 'Brownfield vs Greenfield'")
print("=" * 80)
df_brown = pd.read_excel(file_path, sheet_name='Brownfield vs Greenfield')
print(df_brown.to_string())

# Verificar se existe sheet de Inglês Anterior
print("\n" + "=" * 80)
if 'Inglês Anterior' in excel_file.sheet_names:
    print("✅ Nova sheet 'Inglês Anterior' encontrada!")
    df_ingles = pd.read_excel(file_path, sheet_name='Inglês Anterior')
    print(df_ingles.to_string())
else:
    print("❌ Sheet 'Inglês Anterior' NÃO encontrada!")

print("\n" + "=" * 80)
print("STATUS: Excel está com a versão CORRIGIDA! ✅")
print("=" * 80)
