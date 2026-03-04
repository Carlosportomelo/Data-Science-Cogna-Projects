import pandas as pd

df_geral = pd.read_csv('data/base_leads_160126.csv')
df_matriculados = pd.read_csv('data/matriculados_histórico.csv')

ids_geral = set(df_geral['Record ID'].astype(str))
ids_matriculados = set(df_matriculados['Record ID'].astype(str))

ids_faltando = ids_matriculados - ids_geral
df_matriculados_faltando = df_matriculados[df_matriculados['Record ID'].isin(ids_faltando)]

print(f'Total de registros faltando: {len(ids_faltando):,}')
print(f'\nAmostra dos primeiros 5 registros faltando:')
for idx, row in df_matriculados_faltando.head(5).iterrows():
    print(f'\nRecord ID: {row["Record ID"]}')
    print(f'  Nome: {row.get("Nome do negócio", "N/A")}')
    print(f'  Data criação: {row.get("Data de criação", "N/A")}')
    print(f'  Unidade: {row.get("Unidade Desejada", "N/A")}')
    print(f'  Etapa: {row.get("Etapa do negócio", "N/A")}')

# Salvar esses registros em um arquivo para análise
df_matriculados_faltando.to_csv('outputs/registros_faltando.csv', index=False, encoding='utf-8')
print(f'\n✓ Arquivo com registros faltando salvo em: outputs/registros_faltando.csv')
