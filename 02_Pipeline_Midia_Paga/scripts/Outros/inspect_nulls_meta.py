import pandas as pd

p='data/meta_dataset.csv'
print('Lendo', p)
try:
    df=pd.read_csv(p)
except Exception as e:
    print('ERRO lendo CSV:', e)
    raise

print('\nShape:', df.shape)
print('\nColunas:', list(df.columns))
print('\nPrimeiras linhas:')
print(df.head(10).to_string(index=False))

print('\nContagem de valores nulos por coluna:')
print(df.isna().sum())

# Verificar colunas chave
for col in ['Dia', 'Valor usado (BRL)', 'Nome da campanha', 'Nome da conta']:
    if col in df.columns:
        nnull = df[col].isna().sum()
        print(f"\nColuna '{col}': {nnull} nulos / {len(df)} linhas")

# Mostrar linhas onde muitas colunas são nulas
mask_many_null = df.isna().sum(axis=1) >= (len(df.columns) - 3)
if mask_many_null.any():
    print('\nExemplos de linhas com muitas colunas nulas:')
    print(df[mask_many_null].head(10).to_string(index=False))
else:
    print('\nNenhuma linha com número excessivo de nulos encontrada')

# Para Janeiro/2026, mostrar linhas com Dia parsado como NaT
if 'Dia' in df.columns:
    df['Dia_parsed'] = pd.to_datetime(df['Dia'], errors='coerce')
    jan_mask = (df['Dia_parsed'].dt.year == 2026) & (df['Dia_parsed'].dt.month == 1)
    jan = df[jan_mask]
    print(f"\nLinhas em jan/2026: {len(jan)}")
    if 'Valor usado (BRL)' in df.columns:
        print('Nulls em Valor usado (BRL) dentro de jan/2026:', jan['Valor usado (BRL)'].isna().sum())
        print('\nExemplos de jan/2026 com Valor usado (BRL) nulo:')
        print(jan[jan['Valor usado (BRL)'].isna()].head(10).to_string(index=False))

print('\nDiagnóstico:')
print('- Verifique se o CSV tem separador correto e linhas de cabeçalho extras.')
print('- Verifique se há aspas e quebras de linha dentro de campos que podem quebrar o parsing.')
print('- Para normalizar, rode: df = pd.read_csv(p, engine="python", sep=",", encoding="utf-8")')
