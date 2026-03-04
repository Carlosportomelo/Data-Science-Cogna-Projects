import pandas as pd

df = pd.read_excel('Novos Negócios 05_02 9h40_com_descobertas.xlsx', sheet_name='Preenchido')

col_g = df.columns[6]  
col_h = df.columns[7]  

print('=' * 80)
print('DIAGNOSTICO - Colunas G e H')
print('=' * 80)
print()
print(f'Coluna G (posicao 6): {col_g}')
print(f'Coluna H (posicao 7): {col_h}')
print()

# Contar problemas
nenhum_em_h = 0
diferentes = 0
linhas_problema = []

for i in range(len(df)):
    val_g = df.iloc[i, 6]
    val_h = df.iloc[i, 7]
    
    if str(val_h).strip() == 'Nenhum Valor':
        nenhum_em_h += 1
        linhas_problema.append((i+1, val_g, val_h))
    
    if val_g != val_h:
        diferentes += 1

print(f'Total de linhas: {len(df)}')
print(f'Linhas com "Nenhum Valor" em H: {nenhum_em_h}')
print(f'Linhas onde G != H: {diferentes}')
print()

if nenhum_em_h > 0:
    print('PROBLEMAS ENCONTRADOS - Mostrando primeiras 10:')
    for linha, g, h in linhas_problema[:10]:
        print(f'  Linha {linha}: G={g} | H={h}')
    print()

print('AMOSTRA DOS DADOS:')
for i in range(min(10, len(df))):
    val_g = df.iloc[i, 6]
    val_h = df.iloc[i, 7]
    status = '✓' if val_g == val_h else '✗'
    print(f'{status} L{i+1}: G={val_g} | H={val_h}')
