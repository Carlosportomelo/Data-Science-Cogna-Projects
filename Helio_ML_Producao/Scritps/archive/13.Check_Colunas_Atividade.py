import pandas as pd
import os

# ==============================================================================
# VERIFICADOR DE COLUNAS DE TEMPO/ATIVIDADE
# Objetivo: Encontrar "Data da Última Atividade" para salvar o modelo
# ==============================================================================

CAMINHO_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CAMINHO_LEADS = os.path.join(CAMINHO_BASE, 'Data', 'hubspot_leads.csv')

print(f"Lendo arquivo: {CAMINHO_LEADS}")

try:
    try:
        df = pd.read_csv(CAMINHO_LEADS, encoding='utf-8', nrows=5) # Lê só o cabeçalho
    except:
        df = pd.read_csv(CAMINHO_LEADS, encoding='latin1', nrows=5)
    
    df.columns = df.columns.str.strip()
    
    print("\n🔎 PROCURANDO COLUNAS DE RASTREABILIDADE...")
    
    # Palavras-chave para buscar
    keywords = ['data', 'date', 'atividade', 'activity', 'últim', 'last', 'contato', 'contact']
    
    encontradas = []
    for col in df.columns:
        if any(k in col.lower() for k in keywords):
            encontradas.append(col)
            
    if encontradas:
        print(f"\n✅ Encontramos {len(encontradas)} colunas candidatas:")
        for c in encontradas:
            print(f"   - {c}")
            
        print("\nSUGESTÃO: Se houver algo como 'Data da última atividade' ou 'Last Activity Date',")
        print("podemos usar isso para substituir a 'Densidade' e aumentar drasticamente a precisão.")
    else:
        print("\n❌ Nenhuma coluna de data de atividade encontrada.")
        print("AÇÃO: Solicitar nova exportação do HubSpot incluindo 'Data da última atividade'.")

except Exception as e:
    print(f"Erro ao ler arquivo: {e}")