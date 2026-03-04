import pandas as pd
import os

# --- CONFIGURAÇÃO ---
CAMINHO_RELATIVO = os.path.join('Data', 'hubspot_leads.csv')
if not os.path.exists(CAMINHO_RELATIVO):
    CAMINHO_RELATIVO = os.path.join('..', 'Data', 'hubspot_leads.csv')

print(f"[DEBUG V3] Lendo arquivo (Linguagem Original): {os.path.abspath(CAMINHO_RELATIVO)}")

try:
    try:
        df = pd.read_csv(CAMINHO_RELATIVO, sep=',', encoding='utf-8', on_bad_lines='skip')
    except:
        df = pd.read_csv(CAMINHO_RELATIVO, sep=',', encoding='latin1', on_bad_lines='skip')

    df.columns = df.columns.str.strip()
    
    # Mantendo os nomes originais para fidelidade ao CSV, apenas simplificando a digitação
    # 'Fonte original do tráfego' -> 'Fonte'
    df = df.rename(columns={
        'Fonte original do tráfego': 'Fonte', 
        'Data de criação': 'Data_Criacao'
    })
    
    df['Data_Criacao'] = pd.to_datetime(df['Data_Criacao'], errors='coerce')

    # --- DEFINIÇÃO DA SAFRA (FILTRO DE DATA) ---
    # Usando o período de "Captação Alta" sugerido (Out/24 a Set/25)
    dt_inicio = '2024-10-01'
    dt_fim = '2025-09-30' # Ajuste aqui se o Dash for até hoje
    
    mask = (df['Data_Criacao'] >= dt_inicio) & (df['Data_Criacao'] <= dt_fim)
    df_safra = df[mask].copy()

    print(f"\n=== PERÍODO ANALISADO: {dt_inicio} até {dt_fim} ===")
    print(f"Total de Linhas no Período: {len(df_safra)}")
    print("-" * 50)

    # --- CONTAGEM EXATA (SEM AGRUPAMENTO) ---
    # Aqui veremos os nomes exatamente como estão no banco
    contagem = df_safra['Fonte'].value_counts().reset_index()
    contagem.columns = ['Fonte Original (CSV)', 'Volume']
    
    print(contagem)
    print("-" * 50)

    # --- SIMULAÇÃO DE TOTAIS (MANUAL) ---
    # Agora somamos apenas o que tem certeza
    
    # 1. TENTATIVA DE 'DIGITAL - PAGO' (Dash: ~8.595)
    # Somando explicitamente as linhas do CSV que são pagas
    vol_social_pago = df_safra[df_safra['Fonte'] == 'Social pago'].shape[0]
    vol_pesq_paga   = df_safra[df_safra['Fonte'] == 'Pesquisa paga'].shape[0]
    total_pago_csv  = vol_social_pago + vol_pesq_paga
    
    # 2. TENTATIVA DE 'CADASTRO DA UNIDADE' (Dash: ~8.035)
    vol_offline     = df_safra[df_safra['Fonte'] == 'Fontes off-line'].shape[0]
    
    print("COMPARATIVO FINAL (Buscando o Match):")
    print(f"Social pago ({vol_social_pago}) + Pesquisa paga ({vol_pesq_paga}) = {total_pago_csv}")
    print(f"Fontes off-line = {vol_offline}")

    if abs(total_pago_csv - 8595) < 500:
        print("\n[VEREDITO] BATEU! O período correto é este.")
    else:
        print("\n[VEREDITO] Ainda diferente. O Dashboard deve estar usando outro filtro de data.")

except Exception as e:
    print(f"Erro: {e}")