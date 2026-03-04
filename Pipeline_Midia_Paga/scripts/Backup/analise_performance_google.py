#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script de Análise de Performance - Google Ads (Versão Corrigida)
Lê 'googleads_dataset.csv', ignora 2 linhas.
Salva 5 abas (YoY, Completo, 2023, 2024, 2025) em 'google_dashboard.xlsx'.
"""

import pandas as pd
import sys
from pathlib import Path
import re
import numpy as np

# --- Constantes ---
try:
    BASE_DIR = Path(__file__).parent.parent
except NameError:
    BASE_DIR = Path.cwd()

DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = Path("outputs")

FILE_PATH = Path("data") / "googleads_dataset.csv"
SKIP_ROWS = 2 # Ignora as 2 primeiras linhas (título, período)

# --- Caminhos de Saída ---
OUT_EXCEL_FILE = OUTPUT_DIR / "google_dashboard.xlsx"

# --- Função Utilitária para Números ---
def parse_number(x):
    """Converte valores monetários brasileiros para float"""
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    s = s.replace(" ", "").replace('"', '')
    
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

print("="*80)
print("📊 PROCESSAMENTO DE DADOS - GOOGLE ADS (Versão Corrigida)")
print("="*80)

# --- 1. Carregar Dados ---
print(f"\n📂 Carregando dados de: {FILE_PATH}")
try:
    # Lendo o CSV com separador ','
    df = pd.read_csv(FILE_PATH, skiprows=SKIP_ROWS, encoding='utf-8', sep=',')
    
    if df.empty:
        raise ValueError("O DataFrame está vazio após o carregamento.")
    
    print(f"✅ {len(df)} linhas brutas carregadas")

except FileNotFoundError:
    print(f"❌ ERRO: Arquivo não encontrado em: {FILE_PATH.resolve()}")
    print("   Verifique se o caminho e o nome do arquivo estão corretos.")
    sys.exit(1)
except Exception as e:
    print(f"❌ ERRO ao carregar o arquivo: {e}")
    sys.exit(1)

# --- 2. Normalizar Colunas ---
print("\n🔧 Normalizando colunas...")
df.columns = df.columns.map(lambda c: str(c).strip() if not pd.isna(c) else c)
df.columns = df.columns.str.strip('"')
# Remove caracteres não-ASCII
df.columns = [re.sub(r'[^\x00-\x7F]+', '', col) for col in df.columns]
df.columns = df.columns.str.replace('\\r', '', regex=False).str.replace('\\n', '', regex=False)

print("\n📋 Colunas detectadas (normalizadas):")
print(list(df.columns))

# --- 3. Mapear Colunas ---
print("\n🔧 Mapeando colunas do Google Ads...")

col_mapping = {
    'Campanha': 'Nome_Campanha',
    'Tipo de campanha': 'Tipo_Campanha',
    'Dia': 'Data',
    'Custo': 'Investimento',
    'Conversões': 'Conversoes',
    'Converses': 'Conversoes',
    'Conversoes': 'Conversoes',
    'Custo / conv.': 'CPL'
}

# Aplicar o renomeio
for col_original, col_nova in col_mapping.items():
    if col_original in df.columns:
        df.rename(columns={col_original: col_nova}, inplace=True)

# --- 4. FILTRO: EXCLUIR VERTICAIS NÃO DESEJADAS (Ex: BILINGUAL) ---
print("\n🔧 Aplicando filtro: Excluindo registros com palavras-chave de exclusão (Ex: 'bilingual', 'bilingue')...")
linhas_antes = len(df)

# Palavras-chave de exclusão (case-insensitive)
EXCLUSION_KEYWORDS = ['bilingual', 'bilingue']

# Colunas que provavelmente contêm o nome da campanha ou conta
# No Google Ads, a coluna principal é 'Nome_Campanha' (que vem de 'Campanha')
colunas_para_verificar = [col for col in df.columns if 'campanha' in str(col).lower() or 'conta' in str(col).lower() or 'nome' in str(col).lower()]

mask_exclusao = pd.Series([False] * len(df), index=df.index)

for col in colunas_para_verificar:
    try:
        # Verifica se a coluna contém qualquer uma das palavras-chave
        for keyword in EXCLUSION_KEYWORDS:
            mask_exclusao |= df[col].astype(str).str.contains(
                keyword, 
                case=False, 
                na=False, 
                regex=False
            )
    except Exception as e:
        print(f"  ⚠️ Aviso: Erro ao processar coluna '{col}': {e}")
        continue

# Aplicar filtro (manter apenas linhas SEM as palavras-chave de exclusão)
df = df[~mask_exclusao].copy()

linhas_removidas = linhas_antes - len(df)
print(f"  ✅ Filtro aplicado: {linhas_removidas} linhas removidas (contendo palavras-chave de exclusão)")
print(f"  ✅ Linhas restantes: {len(df)}")

if df.empty:
    print("❌ ERRO: Nenhuma linha restou após a exclusão. Verifique os dados.")
    sys.exit(1)

# --- 5. Validar Colunas ---
colunas_obrigatorias = ['Data', 'Investimento', 'Conversoes', 'Tipo_Campanha']
colunas_faltantes = [col for col in colunas_obrigatorias if col not in df.columns]

if colunas_faltantes:
    print(f"\n❌ ERRO: Colunas obrigatórias não encontradas: {colunas_faltantes}")
    print(f"   Verifique se o arquivo '{FILE_PATH.name}' tem as colunas corretas.")
    sys.exit(1)
else:
    print("   ✅ Colunas essenciais (Data, Investimento, Conversoes, Tipo_Campanha) encontradas.")

# --- 5. Processar Data ---
print(f"\n🔧 Processando coluna de data...")
df['Data_Datetime'] = pd.to_datetime(df['Data'], errors='coerce')
num_na_dates = df['Data_Datetime'].isna().sum()
if num_na_dates > 0:
    print(f"⚠️ {num_na_dates} linhas com data inválida serão removidas.")
    df = df.dropna(subset=['Data_Datetime'])

if df.empty:
    print("❌ ERRO: Nenhuma linha válida após conversão de data.")
    sys.exit(1)

print(f"✅ {len(df)} linhas válidas")

# --- 6. Processar Valores Numéricos ---
print(f"\n🔧 Processando valores numéricos...")
df['Investimento_Google'] = df['Investimento'].apply(parse_number)
df['Leads_Google'] = df['Conversoes'].apply(parse_number)

# --- 7. Adicionar Colunas de Tempo ---
print(f"\n🔧 Adicionando colunas de tempo...")
df['Ano'] = df['Data_Datetime'].dt.year
df['Mes'] = df['Data_Datetime'].dt.month
df['Mes_Ano'] = df['Data_Datetime'].dt.to_period('M')

# --- 8. Adicionar Coluna de Atribuição HubSpot ---
print(f"\n🔧 Adicionando coluna 'Tipo_campanha_HUBSPOT' = 'Pesquisa Paga'")
df['Tipo_campanha_HUBSPOT'] = 'Pesquisa Paga'

print("\n✅ Processamento básico concluído com sucesso!")

# =====================================================================
# --- 9. GERAR RELATÓRIOS ---
# =====================================================================
print("\nGerando relatórios...")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# --- Relatório 1: Google_YoY (Agregado por Dia) ---
print(f"  Processando dados para a aba 'Google_YoY'...")
df_daily_agg = df.groupby('Data_Datetime').agg(
    Investimento_Google=('Investimento_Google', 'sum'),
    Leads_Google=('Leads_Google', 'sum')
).reset_index()

df_daily_agg = df_daily_agg.rename(columns={'Data_Datetime': 'Data'})
# Manter linhas com investimento OU leads para não perder dados
df_daily_agg = df_daily_agg[(df_daily_agg['Investimento_Google'] > 0) | (df_daily_agg['Leads_Google'] > 0)].sort_values(by='Data')

# Adicionando o cálculo de CPL
df_daily_agg['CPL_Google'] = df_daily_agg['Investimento_Google'] / df_daily_agg['Leads_Google']
df_daily_agg['CPL_Google'] = df_daily_agg['CPL_Google'].fillna(0).replace([np.inf, -np.inf], 0)


# --- Relatório 2: Abas por Ano (2023, 2024, 2025) ---
print("  Processando dados para as abas por ano...")
df_2023 = df[df['Ano'] == 2023]
df_2024 = df[df['Ano'] == 2024]
df_2025 = df[df['Ano'] == 2025]

# Salvar tudo em um único arquivo Excel com abas
print(f"\n💾 Salvando arquivo Excel único em: {OUT_EXCEL_FILE}")
try:
    with pd.ExcelWriter(OUT_EXCEL_FILE, engine='openpyxl') as writer:
        # Aba 1: Google_YoY
        df_daily_agg.to_excel(writer, sheet_name='Google_YoY', index=False, float_format='%.2f')
        print("  ✅ Aba 'Google_YoY' salva.")
        
        # Aba 2: Google_Completo (O dataframe 'df' original processado)
        df.to_excel(writer, sheet_name='Google_Completo', index=False)
        print("  ✅ Aba 'Google_Completo' (com Tipo_campanha_HUBSPOT) salva.")

        # Abas por Ano
        if not df_2023.empty:
            df_2023.to_excel(writer, sheet_name='Google_2023', index=False)
            print("  ✅ Aba 'Google_2023' salva.")
        if not df_2024.empty:
            df_2024.to_excel(writer, sheet_name='Google_2024', index=False)
            print("  ✅ Aba 'Google_2024' salva.")
        if not df_2025.empty:
            df_2025.to_excel(writer, sheet_name='Google_2025', index=False)
            print("  ✅ Aba 'Google_2025' salva.")
    
    print(f"\n✅ Arquivo Excel '{OUT_EXCEL_FILE.name}' gerado com sucesso na pasta '{OUTPUT_DIR}'!")

except ImportError:
    print("\n\n❌ ERRO: A BIBLIOTECA 'openpyxl' NÃO ESTÁ INSTALADA.")
    print("Para salvar em Excel, por favor, rode o comando no seu terminal:")
    print("pip install openpyxl")
    sys.exit(1)
except Exception as e:
    print(f"\n\n❌ ERRO AO SALVAR O EXCEL: {e}")
    print("Verifique se o arquivo não está aberto em outro programa.")
    sys.exit(1)


# =====================================================================
# --- 10. CONFIRMAÇÃO DE DADOS ---
# =====================================================================
print("\n--- Confirmação de Investimento Google (2025) ---")

try:
    df_daily_agg['Data'] = pd.to_datetime(df_daily_agg['Data'])
    
    # Calcular Setembro 2025
    invest_set_2025 = df_daily_agg[
        (df_daily_agg['Data'].dt.year == 2025) &
        (df_daily_agg['Data'].dt.month == 9)
    ]['Investimento_Google'].sum()
    
    # Calcular Outubro 2025
    invest_out_2025 = df_daily_agg[
        (df_daily_agg['Data'].dt.year == 2025) &
        (df_daily_agg['Data'].dt.month == 10)
    ]['Investimento_Google'].sum()
    
    # Formatar como moeda brasileira
    print(f"  ✅ Investimento Total em Setembro/2025: R$ {invest_set_2025:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    print(f"  ✅ Investimento Total em Outubro/2025:   R$ {invest_out_2025:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

except Exception as e_conf:
    print(f"  ⚠️ Não foi possível calcular a confirmação de investimento: {e_conf}")

print("\n--- Amostra do Relatório YoY (Aba 'Google_YoY') ---")
print(df_daily_agg.head())

if __name__ == "__main__":
    pass