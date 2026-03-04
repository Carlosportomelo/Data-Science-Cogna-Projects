import pandas as pd
import os
from datetime import datetime

# --- CONFIGURAÇÃO ---
CAMINHO_ENTRADA = os.path.join('Data', 'hubspot_leads.csv')
if not os.path.exists(CAMINHO_ENTRADA):
    CAMINHO_ENTRADA = os.path.join('..', 'Data', 'hubspot_leads.csv')

PASTA_OUTPUT = os.path.join(os.path.dirname(CAMINHO_ENTRADA), '..', 'Outputs')
os.makedirs(PASTA_OUTPUT, exist_ok=True)

DATA_HOJE = datetime.now().strftime('%Y-%m-%d')
# Adicionei "FULL" no nome para você saber que é o completo
CAMINHO_SAIDA = os.path.join(PASTA_OUTPUT, f'Validacao_Rastreabilidade_FULL_{DATA_HOJE}.xlsx')

print(f"[INIT] Gerando Relatório COMPLETO (100k+) em: {CAMINHO_SAIDA}")

try:
    # 1. CARREGAMENTO
    try:
        df = pd.read_csv(CAMINHO_ENTRADA, sep=',', encoding='utf-8', on_bad_lines='skip')
    except:
        df = pd.read_csv(CAMINHO_ENTRADA, sep=',', encoding='latin1', on_bad_lines='skip')

    df.columns = df.columns.str.strip()
    
    mapa = {
        'Record ID': 'ID_Lead',
        'Data de criação': 'Data_Entrada',
        'Fonte original do tráfego': 'Fonte_Original_HubSpot', 
        'Detalhamento da fonte original do tráfego 1': 'Detalhe_Tecnico_Campanha', 
        'Etapa do negócio': 'Status_Atual'
    }
    df = df.rename(columns=mapa)
    df['Data_Entrada'] = pd.to_datetime(df['Data_Entrada'], errors='coerce')

    # --- REMOVI O FILTRO DE DATA AQUI ---
    # Antes tinha: df = df[df['Data_Entrada'] >= '2024-01-01']
    # Agora vai processar tudo.

    # 2. CLASSIFICAÇÃO
    print(f"[PROCESSANDO] Classificando origens de {len(df)} leads...")
    
    def classificar_origem_demo(row):
        detalhe = str(row['Detalhe_Tecnico_Campanha']).lower()
        macro = str(row['Fonte_Original_HubSpot']).lower()
        
        if 'facebook' in detalhe or 'instagram' in detalhe:
            return 'Meta Ads (Social)'
        elif 'google' in detalhe or 'adwords' in detalhe or 'pmax' in detalhe or 'search' in detalhe:
            return 'Google Ads (Busca)'
        elif 'tiktok' in detalhe:
            return 'TikTok'
        elif 'unknown keywords' in detalhe or 'organica' in macro:
            return 'Google Orgânico'
        elif 'off-line' in macro or 'unidade' in detalhe:
            return 'Offline / Balcão'
        elif 'pago' in macro or 'paid' in macro:
            return 'Outra Mídia Paga'
        else:
            return 'Outros'

    df['> CLASSIFICACAO_NOSSA <'] = df.apply(classificar_origem_demo, axis=1)

    # 3. EXPORTAÇÃO
    colunas_demo = [
        'ID_Lead',
        'Data_Entrada',
        'Fonte_Original_HubSpot', 
        'Detalhe_Tecnico_Campanha', 
        '> CLASSIFICACAO_NOSSA <', 
        'Status_Atual'
    ]
    
    colunas_finais = [c for c in colunas_demo if c in df.columns]
    df_demo = df[colunas_finais]

    print(f"[EXPORTANDO] Escrevendo {len(df_demo)} linhas no Excel... (Isso pode demorar um pouco)")
    
    df_demo.to_excel(CAMINHO_SAIDA, index=False, engine='openpyxl', sheet_name='Rastreabilidade_Leads')

    print(f"\n[SUCESSO] Relatório FULL gerado: {CAMINHO_SAIDA}")

except Exception as e:
    print(f"[ERRO] {e}")