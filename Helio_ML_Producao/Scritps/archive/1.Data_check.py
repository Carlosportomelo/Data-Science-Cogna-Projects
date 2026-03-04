import pandas as pd
import os

# --- CONFIGURAÇÃO ---
CAMINHO_RELATIVO = os.path.join('Data', 'hubspot_leads.csv')

# Ajuste de caminho relativo para rodar tanto da raiz quanto da pasta Scripts
if not os.path.exists(CAMINHO_RELATIVO):
    CAMINHO_RELATIVO = os.path.join('..', 'Data', 'hubspot_leads.csv')

print(f"[INIT] Procurando arquivo em: {os.path.abspath(CAMINHO_RELATIVO)}")

try:
    # 1. LEITURA
    if not os.path.exists(CAMINHO_RELATIVO):
        raise FileNotFoundError(f"Arquivo não encontrado: {CAMINHO_RELATIVO}")

    try:
        df = pd.read_csv(CAMINHO_RELATIVO, sep=',', encoding='utf-8')
    except UnicodeDecodeError:
        print("[AVISO] UTF-8 falhou. Tentando latin1...")
        df = pd.read_csv(CAMINHO_RELATIVO, sep=',', encoding='latin1')

    # 2. PADRONIZAÇÃO DE COLUNAS (A Correção está aqui)
    df.columns = df.columns.str.strip() # Remove espaços extras
    
    # Mapeia do nome que ESTÁ no arquivo (minúsculo) para o nome PADRÃO (Título)
    mapa_colunas = {
        'Data de criação': 'Data de Criação',             # Corrigido: c -> C
        'Data de criaÃ§Ã£o': 'Data de Criação',           # Mantido por segurança
        'Nome do negócio': 'Nome do Negócio',             # Corrigido: n -> N
        'Etapa do negócio': 'Etapa do Negócio',           # Corrigido: n -> N
        'Fonte original do tráfego': 'Fonte de Tráfego',  # Simplificando
        'Data de fechamento': 'Data de Fechamento'
    }
    
    # Aplica a renomeação
    df = df.rename(columns=mapa_colunas)

    print("\n--- Colunas Padronizadas ---")
    print(df.columns.tolist())

    # 3. VALIDAÇÃO
    coluna_alvo = 'Data de Criação'
    
    if coluna_alvo in df.columns:
        # Converte para data
        df[coluna_alvo] = pd.to_datetime(df[coluna_alvo], errors='coerce')
        
        # Estatísticas básicas
        print(f"\n[SUCESSO] Dados carregados e processados.")
        print(f"Total de registros: {len(df)}")
        print(f"Datas válidas: {df[coluna_alvo].notna().sum()}")
        
        # Verifica se temos dados recentes (importante para sua análise de recuperação)
        if not df[coluna_alvo].isnull().all():
            print(f"Período: de {df[coluna_alvo].min().date()} até {df[coluna_alvo].max().date()}")
        
        # Check rápido nas colunas categóricas para seus futuros modelos de classificação
        if 'Etapa do Negócio' in df.columns:
            print("\nTop 3 Etapas do Negócio:")
            print(df['Etapa do Negócio'].value_counts().head(3))
            
    else:
        print(f"\n[ERRO] A coluna '{coluna_alvo}' ainda não foi encontrada.")
        print("Verifique o print das 'Colunas Padronizadas' acima.")

except Exception as e:
    print(f"\n[ERRO CRÍTICO] {e}")