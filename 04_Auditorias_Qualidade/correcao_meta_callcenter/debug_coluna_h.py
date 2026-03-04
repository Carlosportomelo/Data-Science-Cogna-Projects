import pandas as pd

df = pd.read_excel('Novos Negócios 05_02 9h40_com_descobertas.xlsx')
print('Total de linhas:', len(df))
print('\n=== Coluna G (Data de Nascimento do Aluno - Diário) ===')
print('Valores vazios:', df.iloc[:, 6].isna().sum())
print('Valores "(Nenhum valor)":', (df.iloc[:, 6] == '(Nenhum valor)').sum())
print('Primeiros 10 valores:')
print(df.iloc[:10, 6])

print('\n=== Coluna H (Data de nascimento do candidato - Diário) ===')
print('Valores vazios:', df.iloc[:, 7].isna().sum())
print('Valores "(Nenhum valor)":', (df.iloc[:, 7] == '(Nenhum valor)').sum())
print('Primeiros 10 valores:')
print(df.iloc[:10, 7])

print('\n=== Resumo ===')
print('Coluna H com "nenhum valor":', (df.iloc[:, 7] == '(Nenhum valor)').sum(), 'de', len(df))

# Tenta trazer dados de coluna G para H onde H está vazio
print('\n=== Tentando preencher H com dados de G ===')
df_copy = df.copy()
col_g = df_copy.iloc[:, 6]
col_h = df_copy.iloc[:, 7]
mask = (col_h == '(Nenhum valor)') | (col_h.isna())
df_copy.iloc[:, 7] = col_h.where(~mask, col_g)
print('Preenchidos:', mask.sum(), 'linhas')
print('\nNovos valores da coluna H (primeiros 10):')
print(df_copy.iloc[:10, 7])

# Salva versão corrigida
df_copy.to_excel('Novos Negócios 05_02 9h40_com_descobertas_CORRIGIDO.xlsx', index=False)
print('\nArquivo salvo como: Novos Negócios 05_02 9h40_com_descobertas_CORRIGIDO.xlsx')
