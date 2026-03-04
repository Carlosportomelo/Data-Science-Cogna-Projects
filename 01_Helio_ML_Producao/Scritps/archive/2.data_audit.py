import pandas as pd
import os

# --- CONFIGURAÇÃO ---
CAMINHO_RELATIVO = os.path.join('Data', 'hubspot_leads.csv')
if not os.path.exists(CAMINHO_RELATIVO):
    CAMINHO_RELATIVO = os.path.join('..', 'Data', 'hubspot_leads.csv')

try:
    # Leitura
    try:
        df = pd.read_csv(CAMINHO_RELATIVO, sep=',', encoding='utf-8', on_bad_lines='skip')
    except:
        df = pd.read_csv(CAMINHO_RELATIVO, sep=',', encoding='latin1', on_bad_lines='skip')
    
    # Padronizar colunas
    df.columns = df.columns.str.strip()
    mapa = {
        'Data de criação': 'Data de Criação', 'Data de criaÃ§Ã£o': 'Data de Criação',
        'Fonte original do tráfego': 'Fonte', 'Nome do negócio': 'Nome',
        'Etapa do negócio': 'Etapa', 'Pipeline': 'Pipeline'
    }
    df = df.rename(columns=mapa)

    print("\n=== 1. QUAIS PIPELINES EXISTEM? (Isso define o recorte) ===")
    if 'Pipeline' in df.columns:
        print(df['Pipeline'].value_counts())
    else:
        print("Coluna 'Pipeline' não encontrada.")

    print("\n=== 2. COMO ESTÁ ESCRITO 'CADASTRO DA UNIDADE'? ===")
    # Vamos ver as 15 fontes mais comuns para achar o nome exato
    if 'Fonte' in df.columns:
        print(df['Fonte'].value_counts().head(15))
    
    print("\n=== 3. TEM ALGUMA 'FONTE' QUE SÓ APARECE EM 2025? ===")
    # Filtrando apenas 2025 para ver se bate com o título do seu Dash "Captação Alta 25-26"
    df['Data de Criação'] = pd.to_datetime(df['Data de Criação'], errors='coerce')
    df_25 = df[df['Data de Criação'].dt.year == 2025]
    print(f"Total de Leads em 2025: {len(df_25)}")
    print("Top 5 Fontes em 2025:")
    print(df_25['Fonte'].value_counts().head(5))

except Exception as e:
    print(e)