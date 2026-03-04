#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script de Análise de Performance - Google Ads (CORRIGIDO)
 IMPLEMENTAÇÃO: Lógica Incremental Inteligente
   - Carrega histórico existente.
   - Substitui dados antigos apenas se houver sobreposição de datas (atualização).
   - Mantém histórico antigo intacto.
"""

import pandas as pd
import sys
from pathlib import Path
import re
import numpy as np
from datetime import datetime

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
SHEET_NAME_HISTORICO = 'Google_Completo'

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
        # Formato brasileiro (1.234,56) -> 1234.56
        s = s.replace(".", "").replace(",", ".")
    else:
        # Tenta tratar (1,234.56) ou (1234,56)
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

# --- Função para Carregar Histórico Existente ---
def carregar_historico_existente(caminho_excel):
    """
    Carrega o dashboard existente se ele existir.
    """
    if not caminho_excel.exists():
        print("Nenhum histórico encontrado. Será criado um novo arquivo.")
        return pd.DataFrame()
    
    try:
        print(f"Carregando histórico existente de: {caminho_excel}")
        df_historico = pd.read_excel(caminho_excel, sheet_name=SHEET_NAME_HISTORICO)
        
        # Garantir que Data_Datetime seja datetime
        if 'Data_Datetime' in df_historico.columns:
            df_historico['Data_Datetime'] = pd.to_datetime(df_historico['Data_Datetime'], errors='coerce')
        
        print(f"   {len(df_historico)} linhas de histórico carregadas")
        return df_historico
    except Exception as e:
        print(f"  Erro ao carregar histórico: {e}")
        print("  Continuando sem histórico (será criado novo arquivo)")
        return pd.DataFrame()

# --- Função para Mesclar Dados (Incremental) ---
def mesclar_dados_incremental(df_historico, df_novos, col_data='Data_Datetime'):
    """
    Mescla dados históricos com novos dados de forma inteligente (Cut-off por data).
    """
    if df_historico.empty:
        print("\nSem histórico prévio. Todos os dados serão considerados novos.")
        return df_novos.copy()
    
    print("\nMesclando dados históricos com novos dados...")
    
    # Garantir datetime e normalizar (remover horas se houver)
    df_historico[col_data] = pd.to_datetime(df_historico[col_data], errors='coerce').dt.normalize()
    df_novos[col_data] = pd.to_datetime(df_novos[col_data], errors='coerce').dt.normalize()
    
    # Estatísticas antes da mesclagem
    linhas_historico = len(df_historico)
    linhas_novas = len(df_novos)
    
    # LÓGICA DE CORTE:
    # Descobrimos a data mais antiga dos NOVOS dados.
    # Mantemos o histórico apenas ANTES dessa data.
    # Isso garante que se baixarmos dados corrigidos de dias anteriores, eles substituam os antigos.
    
    data_min_novos = df_novos[col_data].min()
    
    if pd.isna(data_min_novos):
        print("Aviso: Novos dados não possuem data válida. Mantendo histórico original.")
        return df_historico

    # Filtra histórico para manter apenas o que é estritamente mais antigo que a nova carga
    df_historico_antigos = df_historico[df_historico[col_data] < data_min_novos].copy()
    
    # Concatena
    df_merged = pd.concat([df_historico_antigos, df_novos], ignore_index=True)
    
    # Ordenar
    df_merged = df_merged.sort_values(by=col_data).reset_index(drop=True)
    
    # Estatísticas
    linhas_finais = len(df_merged)
    linhas_antigos_mantidos = len(df_historico_antigos)
    linhas_substituidas = linhas_historico - linhas_antigos_mantidos
    
    print(f"  Data de corte (início dos novos dados): {data_min_novos.strftime('%d/%m/%Y')}")
    print(f"  Histórico mantido (anterior ao corte): {linhas_antigos_mantidos} linhas")
    print(f"  Histórico substituído (sobreposto): {linhas_substituidas} linhas")
    print(f"  Novos dados adicionados: {linhas_novas} linhas")
    print(f"  Total final: {linhas_finais} linhas")
    
    return df_merged

# =====================================================================
# --- INÍCIO DO PROCESSAMENTO ---
# =====================================================================

print("="*80)
print("SISTEMA DE PROCESSAMENTO INCREMENTAL - GOOGLE ADS")
print("="*80)

# --- 0. Carregar Histórico ---
df_historico = carregar_historico_existente(OUT_EXCEL_FILE)

# --- 1. Carregar Dados Novos ---
print(f"\nCarregando dados de: {FILE_PATH}")
try:
    # Lendo o CSV com separador ','
    df_novos = pd.read_csv(FILE_PATH, skiprows=SKIP_ROWS, encoding='utf-8', sep=',')
    
    if df_novos.empty:
        raise ValueError("O DataFrame está vazio após o carregamento.")
    
    print(f" {len(df_novos)} linhas carregadas (novos dados)")

except FileNotFoundError:
    print(f" ERRO: Arquivo não encontrado em: {FILE_PATH.resolve()}")
    sys.exit(1)
except Exception as e:
    print(f" ERRO ao carregar o arquivo: {e}")
    sys.exit(1)

# --- 2. Normalizar Colunas ---
print("\nNormalizando colunas...")
df_novos.columns = df_novos.columns.map(lambda c: str(c).strip() if not pd.isna(c) else c)
df_novos.columns = df_novos.columns.str.strip('"')
# Remove caracteres não-ASCII
df_novos.columns = [re.sub(r'[^\x00-\x7F]+', '', col) for col in df_novos.columns]
df_novos.columns = df_novos.columns.str.replace('\\r', '', regex=False).str.replace('\\n', '', regex=False)

print("\nColunas detectadas (normalizadas):")
print(list(df_novos.columns))

# --- 3. Mapear Colunas ---
print("\nMapeando colunas do Google Ads...")

col_mapping = {
    'Campanha': 'Nome_Campanha',
    'Tipo de campanha': 'Tipo_Campanha',
    'Dia': 'Data',
    'Custo': 'Investimento',
    'Conversões (por momento da conv.)': 'Conversoes',
    'Converses (por momento da conv.)': 'Conversoes',
    'Conversões': 'Conversoes',
    'Converses': 'Conversoes',
    'Custo / conv.': 'CPL'
}

# Aplicar o renomeio
for col_original, col_nova in col_mapping.items():
    if col_original in df_novos.columns:
        df_novos.rename(columns={col_original: col_nova}, inplace=True)
        if col_nova == 'Conversoes':
            print(f"   Coluna de conversões mapeada: '{col_original}' → '{col_nova}'")

# --- 4. FILTRO: EXCLUIR VERTICAIS NÃO DESEJADAS ---
print("\nAplicando filtro: Excluindo registros com 'bilingual' ou 'bilingue'...")
linhas_antes = len(df_novos)

EXCLUSION_KEYWORDS = ['bilingual', 'bilingue']

colunas_para_verificar = [col for col in df_novos.columns if 'campanha' in str(col).lower() or 'conta' in str(col).lower() or 'nome' in str(col).lower()]

mask_exclusao = pd.Series([False] * len(df_novos), index=df_novos.index)

for col in colunas_para_verificar:
    try:
        for keyword in EXCLUSION_KEYWORDS:
            mask_exclusao |= df_novos[col].astype(str).str.contains(
                keyword, 
                case=False, 
                na=False, 
                regex=False
            )
    except Exception as e:
        print(f"  Aviso: Erro ao processar coluna '{col}': {e}")
        continue

df_novos = df_novos[~mask_exclusao].copy()

linhas_removidas = linhas_antes - len(df_novos)
print(f"  Filtro aplicado: {linhas_removidas} linhas removidas")
print(f"  Linhas restantes: {len(df_novos)}")

if df_novos.empty:
    print(" ERRO: Nenhuma linha restou após a exclusão.")
    sys.exit(1)

# --- 5. Validar Colunas ---
colunas_obrigatorias = ['Data', 'Investimento', 'Conversoes', 'Tipo_Campanha']
colunas_faltantes = [col for col in colunas_obrigatorias if col not in df_novos.columns]

if colunas_faltantes:
    print(f"\nERRO: Colunas obrigatórias não encontradas: {colunas_faltantes}")
    sys.exit(1)
else:
    print("   Colunas essenciais encontradas.")

# --- 6. Processar Data ---
print(f"\nProcessando coluna de data...")
df_novos['Data_Datetime'] = pd.to_datetime(df_novos['Data'], errors='coerce')
num_na_dates = df_novos['Data_Datetime'].isna().sum()
if num_na_dates > 0:
    print(f" {num_na_dates} linhas com data inválida serão removidas.")
    df_novos = df_novos.dropna(subset=['Data_Datetime'])

if df_novos.empty:
    print(" ERRO: Nenhuma linha válida após conversão de data.")
    sys.exit(1)

print(f" {len(df_novos)} linhas válidas")

# --- 7. Processar Valores Numéricos ---
print(f"\nProcessando valores numéricos...")
df_novos['Investimento_Google'] = df_novos['Investimento'].apply(parse_number)
df_novos['Conversoes_Google'] = df_novos['Conversoes'].apply(parse_number)

# Garantir que são realmente números
df_novos['Investimento_Google'] = pd.to_numeric(df_novos['Investimento_Google'], errors='coerce').fillna(0)
df_novos['Conversoes_Google'] = pd.to_numeric(df_novos['Conversoes_Google'], errors='coerce').fillna(0)

print(f"   Investimento convertido para numérico")
print(f"   Conversões convertidas para numérico")

# --- 8. Adicionar Colunas de Tempo ---
print(f"\nAdicionando colunas de tempo...")
df_novos['Ano'] = df_novos['Data_Datetime'].dt.year
df_novos['Mes'] = df_novos['Data_Datetime'].dt.month
df_novos['Mes_Ano'] = df_novos['Data_Datetime'].dt.to_period('M')

# --- 9. Adicionar Coluna de Atribuição HubSpot ---
print(f"\nAdicionando coluna 'Tipo_campanha_HUBSPOT' = 'Pesquisa Paga'")
df_novos['Tipo_campanha_HUBSPOT'] = 'Pesquisa Paga'

print("\nProcessamento dos novos dados concluído!")

# --- 10. MESCLAR COM HISTÓRICO (INCREMENTAL) ---
df = mesclar_dados_incremental(df_historico, df_novos, col_data='Data_Datetime')

# --- 11. REMOVER DUPLICATAS FINAIS (Segurança) ---
print("\nVerificando duplicatas exatas...")
linhas_antes_dedup = len(df)
df = df.drop_duplicates(subset=['Nome_Campanha', 'Data_Datetime', 'Tipo_Campanha'], keep='last')
linhas_depois_dedup = len(df)
linhas_removidas_dedup = linhas_antes_dedup - linhas_depois_dedup

print(f"  Linhas antes: {linhas_antes_dedup}")
print(f"  Duplicatas removidas: {linhas_removidas_dedup}")
print(f"  Linhas após deduplicação: {linhas_depois_dedup}")

# Reordenar por data
df = df.sort_values(by='Data_Datetime').reset_index(drop=True)

print("Deduplicação concluída!")

# =====================================================================
# --- 12. GERAR RELATÓRIOS ---
# =====================================================================
print("\nGerando relatórios...")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# --- Relatório 1: Google_YoY (Agregado por Dia) ---
print(f"  Processando dados para a aba 'Google_YoY'...")

df_daily_agg = df.groupby('Data_Datetime').agg(
    Investimento_Google=('Investimento_Google', 'sum'),
    Conversoes_Google=('Conversoes_Google', 'sum')
).reset_index()

df_daily_agg = df_daily_agg.rename(columns={'Data_Datetime': 'Data'})

# Manter linhas com investimento OU conversões
df_daily_agg = df_daily_agg[(df_daily_agg['Investimento_Google'] > 0) | (df_daily_agg['Conversoes_Google'] > 0)].sort_values(by='Data')

# Adicionando o cálculo de CPL
df_daily_agg['CPL_Google'] = df_daily_agg['Investimento_Google'] / df_daily_agg['Conversoes_Google']
df_daily_agg['CPL_Google'] = df_daily_agg['CPL_Google'].fillna(0).replace([np.inf, -np.inf], 0)

# Arredondar valores
df_daily_agg['Investimento_Google'] = df_daily_agg['Investimento_Google'].round(2).astype(float)
df_daily_agg['Conversoes_Google'] = df_daily_agg['Conversoes_Google'].round(0).astype(int)
df_daily_agg['CPL_Google'] = df_daily_agg['CPL_Google'].round(2).astype(float)

# --- Relatório 2: Abas por Ano (2023, 2024, 2025) ---
print("  Processando dados para as abas por ano...")

# Criar dicionário de dataframes por ano dinamicamente
anos_disponiveis = sorted(df['Ano'].dropna().unique().astype(int))
dfs_por_ano = {ano: df[df['Ano'] == ano].copy() for ano in anos_disponiveis}

# Salvar tudo em um único arquivo Excel com abas
print(f"\nSalvando arquivo Excel em: {OUT_EXCEL_FILE}")
try:
    with pd.ExcelWriter(OUT_EXCEL_FILE, engine='openpyxl') as writer:
        # Aba 1: Google_YoY
        df_daily_agg.to_excel(writer, sheet_name='Google_YoY', index=False)
        print("  Aba 'Google_YoY' salva.")
        
        # Aba 2: Google_Completo
        df.to_excel(writer, sheet_name='Google_Completo', index=False)
        print("  Aba 'Google_Completo' salva.")

        # Abas por Ano
        for ano, df_ano in dfs_por_ano.items():
            nome_aba = f'Google_{ano}'
            if not df_ano.empty:
                df_ano.to_excel(writer, sheet_name=nome_aba, index=False)
                print(f"  Aba '{nome_aba}' salva com {len(df_ano)} registros.")
    
    print(f"\nArquivo Excel '{OUT_EXCEL_FILE.name}' gerado com sucesso!")

except Exception as e:
    print(f"\nERRO AO SALVAR O EXCEL PRINCIPAL: {e}")
    # Não sai imediatamente para tentar salvar o resumido

# --- GERAÇÃO DO ARQUIVO RESUMIDO PARA DASHBOARD ---
RESUMIDO_DIR = OUTPUT_DIR / "visão resumida"
RESUMIDO_DIR.mkdir(parents=True, exist_ok=True)
OUT_RESUMIDO = RESUMIDO_DIR / "google_resumido_Dash.xlsx"
print(f"\nSalvando arquivo resumido em: {OUT_RESUMIDO}")
try:
    with pd.ExcelWriter(OUT_RESUMIDO, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Google_Completo', index=False)
    print("  Arquivo resumido salvo (Aba: Google_Completo).")
except Exception as e:
    print(f"ERRO AO SALVAR O EXCEL RESUMIDO: {e}")

# =====================================================================
# --- 13. CONFIRMAÇÃO DE DADOS ---
# =====================================================================
print("\n" + "="*80)
print("ESTATÍSTICAS FINAIS - GOOGLE ADS")
print("="*80)

print(f"\nTotal de registros: {len(df)}")
print(f"Período: {df['Data_Datetime'].min().strftime('%d/%m/%Y')} até {df['Data_Datetime'].max().strftime('%d/%m/%Y')}")
print(f"\nInvestimento total: R$ {df['Investimento_Google'].sum():,.2f}")
print(f"Conversões total: {df['Conversoes_Google'].sum():,.0f}")

# Por ano
for ano in sorted(df['Ano'].unique()):
    df_ano = df[df['Ano'] == ano]
    print(f"\n{ano}:")
    print(f"  Registros: {len(df_ano)}")
    print(f"  Investimento: R$ {df_ano['Investimento_Google'].sum():,.2f}")
    print(f"  Conversões: {df_ano['Conversoes_Google'].sum():,.0f}")

print("\n" + "="*80)
print("PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
print("="*80)