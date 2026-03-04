import pandas as pd
import numpy as np
import os
import sys

# ==============================================================================
# 1. CONFIGURAÇÃO DE AMBIENTE E CAMINHOS
# ==============================================================================
# Define o caminho do arquivo baseando-se na estrutura que você me passou
caminho_csv = os.path.join('Data', 'hubspot_leads.csv')

# Lógica de segurança: se rodar de dentro da pasta Scripts, sobe um nível
if not os.path.exists(caminho_csv):
    caminho_csv = os.path.join('..', 'Data', 'hubspot_leads.csv')

# Cria pasta de Outputs se não existir (para salvar os relatórios)
pasta_output = os.path.join(os.path.dirname(caminho_csv), '..', 'Outputs')
os.makedirs(pasta_output, exist_ok=True)

print(f"[INIT] Lendo base de dados: {os.path.abspath(caminho_csv)}")

# ==============================================================================
# 2. CARREGAMENTO E LIMPEZA INICIAL (Ajustes de Auditoria)
# ==============================================================================
try:
    # Tenta ler UTF-8, se falhar vai para Latin1 (padrão Excel Brasil)
    try:
        df = pd.read_csv(caminho_csv, sep=',', encoding='utf-8', on_bad_lines='skip')
    except UnicodeDecodeError:
        df = pd.read_csv(caminho_csv, sep=',', encoding='latin1', on_bad_lines='skip')

    # Remove espaços em branco dos nomes das colunas
    df.columns = df.columns.str.strip()

    # Dicionário de Tradução: Do "HubSpot Técnico" para "Português Python"
    mapa_colunas = {
        'Data de criação': 'Data de Criação', 
        'Data de criaÃ§Ã£o': 'Data de Criação', # Prevenção contra erro de encoding
        'Nome do negócio': 'Nome do Negócio',
        'Etapa do negócio': 'Etapa do Negócio',
        'Fonte original do tráfego': 'Fonte Macro',
        'Detalhamento da fonte original do tráfego 1': 'Detalhe da Fonte',
        'Data de fechamento': 'Data de Fechamento'
    }
    df = df.rename(columns=mapa_colunas)
    
    # Conversão de Data
    df['Data de Criação'] = pd.to_datetime(df['Data de Criação'], errors='coerce')

    print(f"[STATUS] Dados carregados. Total de linhas: {len(df)}")

except Exception as e:
    print(f"[ERRO CRÍTICO] Não foi possível ler o arquivo. Detalhes: {e}")
    sys.exit()

# ==============================================================================
# 3. ENGENHARIA DE FEATURES (A Inteligência do Modelo)
# ==============================================================================
# Aqui aplicamos a lógica que definimos para separar Google de Meta

def classificar_canal_inteligente(linha):
    """
    Função que lê os detalhes técnicos da URL/Campanha e define
    se é Google (Busca), Meta (Social), Orgânico ou Outros.
    """
    # Converte para texto minúsculo para facilitar a busca (case insensitive)
    detalhe = str(linha['Detalhe da Fonte']).lower()
    fonte_macro = str(linha['Fonte Macro']).lower()

    # --- LÓGICA DE CLASSIFICAÇÃO ---
    
    # 1. META ADS (Facebook/Instagram)
    # Identificamos pelo nome da plataforma no detalhe
    if 'facebook' in detalhe or 'instagram' in detalhe:
        return 'Meta Ads (Social)'
    
    # 2. GOOGLE ADS (Alta Intenção/Busca)
    # Identificamos por termos comuns de campanhas de performance
    elif ('google' in detalhe or 'adwords' in detalhe or 
          'pmax' in detalhe or 'search' in detalhe):
        return 'Google Ads (Search/PMax)'
    
    # 3. VÍDEO (TikTok/YouTube)
    elif 'tiktok' in detalhe:
        return 'TikTok Ads'
    elif 'youtube' in detalhe:
        return 'YouTube Ads'

    # 4. ORGÂNICO (SEO/Google Gratuito)
    # 'unknown keywords' é o padrão do Google para busca segura (HTTPS)
    elif 'unknown keywords' in detalhe or 'organica' in fonte_macro or 'orgânica' in fonte_macro:
        return 'Google Orgânico (SEO)'

    # 5. OFFLINE / CADASTRO DIRETO
    # Geralmente balcão, telefone ou integração manual
    elif 'offline' in fonte_macro or 'unidade' in detalhe:
        return 'Offline / Balcão'

    # 6. OUTROS (Email, Referência, Direto sem tracking)
    else:
        return 'Outros / Tráfego Direto'

# Aplica a função linha a linha
print("[PROCESSANDO] Classificando canais de tráfego...")
df['Canal_Modelo'] = df.apply(classificar_canal_inteligente, axis=1)

# ==============================================================================
# 4. DEFINIÇÃO DE SUCESSO (TARGET)
# ==============================================================================
# Cria uma coluna binária (0 ou 1) para usar em Regressão Logística depois
# 1 = Matrícula Realizada, 0 = Qualquer outra coisa
df['Conversao'] = df['Etapa do Negócio'].astype(str).str.contains('MATRÍCULA', case=False).astype(int)

# ==============================================================================
# 5. ANÁLISE DE PERFORMANCE E CORRELAÇÃO
# ==============================================================================
print("[PROCESSANDO] Calculando taxas de conversão por canal...")

# Agrupa os dados pelo novo canal inteligente
analise_canais = df.groupby('Canal_Modelo').agg(
    Leads_Totais=('Nome do Negócio', 'count'),
    Matriculas=('Conversao', 'sum')
).reset_index()

# Calcula a Taxa de Conversão (%)
analise_canais['Taxa_Conversao'] = (analise_canais['Matriculas'] / analise_canais['Leads_Totais'] * 100).round(2)

# Ordena pelo volume de matrículas (quem traz mais dinheiro primeiro)
analise_canais = analise_canais.sort_values(by='Matriculas', ascending=False)

# ==============================================================================
# 6. EXPORTAÇÃO DOS RESULTADOS
# ==============================================================================
arquivo_saida = os.path.join(pasta_output, 'analise_leads_por_fonte.csv')
analise_canais.to_csv(arquivo_saida, index=False, sep=';', encoding='utf-8-sig')

print("\n" + "="*50)
print("RESUMO DA PERFORMANCE (Top Canais)")
print("="*50)
print(analise_canais.to_string(index=False))
print("\n")
print(f"[SUCESSO] Relatório detalhado salvo em: {arquivo_saida}")
print("="*50)