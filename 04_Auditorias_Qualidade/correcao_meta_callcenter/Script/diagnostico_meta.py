from saneamento_leads import find_hubspot_file, read_hubspot, read_meta_all
import pandas as pd
from saneamento_leads import BASE_DIR, DATA_DIR, META_DIR

hub_file = find_hubspot_file(DATA_DIR)
df_hs = read_hubspot(hub_file)
df_meta = read_meta_all(META_DIR)

# Detect target column same as amostra_sucesso logic
col_patterns = ['data de nascimento do aluno', 'data de nascimento', 'nascimento', 'data_nasc', 'diário', 'diario']
target = None
for c in df_hs.columns:
    if any(p in str(c).lower() for p in col_patterns):
        target = c
        break

print(f"Coluna alvo detectada: {target}")

# identify filtered leads (original empty)
def is_original_empty(val):
    if pd.isna(val):
        return True
    s = str(val).strip().lower()
    return s == '' or 'nenhum' in s or 'nenhuma' in s or s in ['(nenhum valor)','nenhum valor','(nenhuma valor)','nenhuma valor']

filtered = df_hs[df_hs[target].apply(is_original_empty)].copy()
print(f"Leads com coluna alvo vazia: {len(filtered)}")

# Merge
merged = pd.merge(filtered, df_meta, on='key_email', how='left')

# counts
has_dob = merged['data_nascimento_meta'].notna().sum()
has_idade_meta = merged['idade_meta'].notna().sum()
# cases with idade but no dob
idade_only = merged[merged['data_nascimento_meta'].isna() & merged['idade_meta'].notna()]
both = merged[merged['data_nascimento_meta'].notna() & merged['idade_meta'].notna()].shape[0]
neither = merged[merged['data_nascimento_meta'].isna() & merged['idade_meta'].isna()].shape[0]

print('\n=== CONCLUSÃO ===')
print(f"Total de leads analisados (com coluna alvo vazia): {len(merged)}")
print(f"1. Leads com Data de Nascimento preenchida: {has_dob}")
print(f"2. Leads com Somente Idade (sem Data): {len(idade_only)}")
print(f"3. Leads com Ambos (Data + Idade): {both}")
print(f"4. Leads sem nenhuma informação: {neither}")

print('\nExemplos (Somente Idade):')
print(idade_only[['key_email','idade_meta','data_nascimento_meta','source_file','campaign']].head(10).to_string(index=False))

print('\nExemplos (Data Disponível):')
print(merged[merged['data_nascimento_meta'].notna()][['key_email','data_nascimento_meta','idade_meta','source_file','campaign']].head(10).to_string(index=False))
