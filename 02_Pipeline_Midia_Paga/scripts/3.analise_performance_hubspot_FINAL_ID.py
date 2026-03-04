#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
analise_performance_hubspot_FINAL_ID.py

 IMPLEMENTAÇÃO: ID Original do HubSpot
   - Usa o Record ID (ou Deal ID) do HubSpot como chave primária (Lead_Key).
"""

import pandas as pd
import numpy as np
import unicodedata
import re
import sys
import hashlib
from pathlib import Path
from datetime import datetime

print("======================================================")
print("Iniciando Script de BLEND - VERSÃO FINAL (ID HUBSPOT)")
print("======================================================")

# --- 1. CONFIGURAÇÕES ---

# Define o BASE_DIR como o PAI da pasta do script
try:
    BASE_DIR = Path(__file__).resolve().parent.parent
except NameError:
    BASE_DIR = Path.cwd()

DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "outputs"
SCRIPT_DIR = BASE_DIR / "scripts"

# Arquivos de entrada (CSV ou XLSX)
HUBSPOT_FILE = Path(r'C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_atual.csv')
# Fallback: Se não encontrar arquivo centralizado, usa pasta local
if not HUBSPOT_FILE.exists():
    if (DATA_DIR / "hubspot_leads.csv").exists():
        HUBSPOT_FILE = DATA_DIR / "hubspot_leads.csv"
    elif (DATA_DIR / "hubspot_dataset.csv").exists():
        HUBSPOT_FILE = DATA_DIR / "hubspot_dataset.csv"
    elif (DATA_DIR / "hubspot_dataset.xlsx").exists():
        HUBSPOT_FILE = DATA_DIR / "hubspot_dataset.xlsx"

META_REPORT_FILE = OUTPUT_DIR / "meta_dataset_dashboard.xlsx" # Assume que os scripts Meta e Google já rodaram
GOOGLE_REPORT_FILE = OUTPUT_DIR / "google_dashboard.xlsx" # Assume que os scripts Meta e Google já rodaram

# Arquivo de saída
BLEND_BASE_NAME = "meta_googleads_blend"

# Configurações de Nome de Conta (Simplificado)
META_ACCOUNT_OTHER_LABEL = "Red Balloon - Contas Meta"
GOOGLE_ACCOUNT_LABEL = "Google Ads"
AREA_GESTAO_DEFAULT = "Gestão Antiga" 

# Regras de Negócio
CANAL_FILTRO = [
    "social pago", 
    "pesquisa paga",
    "paid social",
    "facebook",
    "instagram",
    "linkedin",
    "cpc",
    "paid search",
    "google ads"
]
MATRICULA_KEYWORDS = ["matricula concluida"]
DEFAULT_NA_TEXT = "Não Mapeado"

# Mapeamento dos Nomes de Canal Finais
CANAL_MAP_FINAL = {
    'social pago': 'Social Pago',
    'facebook': 'Social Pago',
    'instagram': 'Social Pago',
    'linkedin': 'Social Pago',
    'paid social': 'Social Pago',
    
    'pesquisa paga': 'Pesquisa Paga',
    'cpc': 'Pesquisa Paga',
    'paid search': 'Pesquisa Paga',
    'google ads': 'Pesquisa Paga'
}

# Mapeamento da Numeração do Funil
ETAPA_FUNIL_MAP = {
    "NOVO NEGÓCIO": "1. Novo Negócio",
    "NEGÓCIO EM QUALIFICAÇÃO": "2. Negócio em Qualificação",
    "VISITA AGENDADA": "3. Visita Agendada",
    "VISITA REALIZADA": "4. Visita Realizada",
    "LISTA DE ESPERA": "5. Lista de Espera",
    "NEGÓCIO EM PAUSA": "6. Negócio em Pausa",
    "NEGÓCIO PERDIDO": "7. Negócio Perdido",
    "MATRÍCULA CONCLUÍDA": "8. Matrícula Realizada"
}
MATRICULA_NOME_FINAL = "8. Matrícula Realizada"


# --- 2. FUNÇÕES UTILITÁRIAS ---

def read_any(path: Path, sheet_name=0, skiprows=0) -> pd.DataFrame:
    """Lê um arquivo CSV ou Excel, com tratamento de erros e múltiplos encodings."""
    
    print(f"   ... Buscando arquivo: {path.resolve()}")
    if not path.exists():
        print(f"\nERRO FATAL: Arquivo não encontrado.")
        print(f"   Caminho buscado: {path.resolve()}")
        sys.exit(1)
    
    suf = path.suffix.lower()
    try:
        if suf in (".xls", ".xlsx"):
            return pd.read_excel(path, sheet_name=sheet_name, skiprows=skiprows)
        elif suf == ".csv" or not suf:
            # Tenta múltiplos encodings
            encodings = ['utf-8', 'latin1', 'iso-8859-1', 'cp1252']
            for encoding in encodings:
                try:
                    df = pd.read_csv(path, engine="python", sep=None, encoding=encoding, skiprows=skiprows)
                    print(f"   ... Arquivo lido com encoding: {encoding}")
                    return df
                except UnicodeDecodeError:
                    continue
            
            # Se nenhum encoding funcionou
            raise Exception("Não foi possível decodificar o arquivo com os encodings testados")
        else:
            print(f"ERRO FINAL: Extensão de arquivo '{suf}' não suportada. Esperado .csv ou .xlsx.")
            sys.exit(1)
            
    except Exception as e:
        print(f"\nERRO FINAL: Não foi possível ler o arquivo {path.name}.")
        print(f"   Detalhe: {e}")
        sys.exit(1)

def clean_cols(df):
    """Limpa e normaliza os nomes das colunas de um DataFrame."""
    cols = []
    for c in df.columns:
        c_str = str(c).strip().lower()
        c_norm = unicodedata.normalize('NFKD', c_str)
        c_ascii = c_norm.encode('ascii', 'ignore').decode('utf-8')
        c_clean = re.sub(r'[^a-z0-9_ ]+', '', c_ascii)
        c_clean = re.sub(r'\s+', '_', c_clean)
        cols.append(c_clean)
    df.columns = cols
    return df

def clean_text(text):
    """Limpa, normaliza e remove acentos de uma string de dados."""
    try:
        text_str = str(text).strip().lower()
        text_norm = unicodedata.normalize('NFKD', text_str)
        text_ascii = text_norm.encode('ascii', 'ignore').decode('utf-8')
        text_clean = re.sub(r'[^a-z0-9 ]+', '', text_ascii)
        return text_clean
    except Exception:
        return ""

def find_col(df: pd.DataFrame, keywords: list, use_clean_cols=False) -> str:
    """
    Encontra a primeira coluna no DataFrame que corresponde a uma keyword.
    """
    target_cols = df.columns
    
    # 1. Busca por correspondência exata
    for k in keywords:
        if k in target_cols:
            return k
            
    # 2. Busca ignorando maiúsculas/minúsculas
    cols_map = {str(c).strip().lower(): str(c) for c in target_cols}
    for k in keywords:
        k_low = k.strip().lower()
        if k_low in cols_map:
            return cols_map[k_low]
            
    # 3. Heurística de substring
    for col_orig in target_cols:
        col_low = str(col_orig).strip().lower()
        for k in keywords:
            if k.strip().lower() in col_low:
                return str(col_orig)
    
    return None

# --- FUNÇÃO DE CÁLCULO DO CICLO DE CAPTAÇÃO ---
def calcular_ciclo_captacao(date_series: pd.Series) -> pd.Series:
    """Calcula o ciclo de captação (YY.1 Alta ou YY.2 Baixa) com base na regra de Outubro-Março."""
    
    df_temp = pd.DataFrame({'date': date_series.copy()})
    
    # Criar série padrão com "Não Mapeado"
    final_series = pd.Series(DEFAULT_NA_TEXT, index=date_series.index)
    
    # Filtrar apenas datas válidas
    valid_mask = df_temp['date'].notna()
    if not valid_mask.any():
        return final_series
    
    df_valid = df_temp[valid_mask].copy()
    df_valid['date'] = pd.to_datetime(df_valid['date'], errors='coerce')
    df_valid = df_valid.dropna(subset=['date'])
    
    if df_valid.empty:
        return final_series
        
    df_valid['ano'] = df_valid['date'].dt.year
    df_valid['mes'] = df_valid['date'].dt.month
    
    # Aplicar lógica de ciclo
    condicoes = [
        (df_valid['mes'] >= 10),  # Outubro, Novembro, Dezembro -> próximo ano .1
        (df_valid['mes'] <= 3),   # Janeiro, Fevereiro, Março -> ano atual .1
        (df_valid['mes'].between(4, 9))  # Abril a Setembro -> ano atual .2
    ]
    
    ciclos = [
        (df_valid['ano'] + 1).astype(str).str[2:] + '.1 Alta',  # Out-Dez
        df_valid['ano'].astype(str).str[2:] + '.1 Alta',        # Jan-Mar
        df_valid['ano'].astype(str).str[2:] + '.2 Baixa'        # Abr-Set
    ]
    
    results = pd.Series(np.select(condicoes, ciclos, default=DEFAULT_NA_TEXT), index=df_valid.index)
    final_series.loc[results.index] = results.values
    
    return final_series
# --- FIM FUNÇÃO DE CÁLCULO DO CICLO DE CAPTAÇÃO ---


# --- FUNÇÃO: EXIBIR MATRÍCULAS CONCLUÍDAS NO TERMINAL ---

def mostrar_matriculas_concluidas(df_mat: pd.DataFrame) -> None:
    """Exibe no terminal os negócios na etapa Matrícula Concluída, ancorados na Data de Fechamento."""
    total = len(df_mat)

    print('\n' + '='*95)
    print('NEGÓCIOS — MATRÍCULA CONCLUÍDA (ancorado em Data de Fechamento)')
    print('='*95)
    print(f"Total de registros: {total}\n")

    if total == 0:
        print("   Nenhum negócio encontrado.")
        print('='*95)
        return

    print(f"{'Data Fechamento':<16} | {'Unidade':<32} | {'Canal':<15} | {'Ciclo':<12} | Lead_Key")
    print('-'*95)

    for _, row in df_mat.iterrows():
        dt = row.get('Data_Fechamento', pd.NaT)
        try:
            data_str = pd.to_datetime(dt, errors='coerce').strftime('%d/%m/%Y')
        except Exception:
            data_str = '—'
        unidade  = str(row.get('Unidade', '—'))[:32]
        canal    = str(row.get('Origem_Principal', '—'))[:15]
        ciclo    = str(row.get('Ciclo_Captacao', '—'))[:12]
        lead_key = str(row.get('Lead_Key', '—'))
        print(f"{data_str:<16} | {unidade:<32} | {canal:<15} | {ciclo:<12} | {lead_key}")

    print('='*95)


# --- 3. PIPELINE PRINCIPAL ---

def main():
    
    print(f"   ... Diretório Base do Projeto: {BASE_DIR.resolve()}")

    # --- 3.1 Carregar e Preparar HubSpot ---
    
    print("\nCarregando dados do HubSpot...")
    df_hub = read_any(HUBSPOT_FILE)
    df_hub = clean_cols(df_hub)
    
    # Colunas esperadas após clean_cols():
    required_hub_cols = [
        'data_de_criacao', 'fonte_original_do_trafego', 'pipeline',
        'detalhamento_da_fonte_original_do_trafego_1', 'detalhamento_da_fonte_original_do_trafego_2',
        'etapa_do_negocio', 'unidade_desejada', 'valor_na_moeda_da_empresa', 'data_de_fechamento'
    ]
    
    missing_cols = [col for col in required_hub_cols if col not in df_hub.columns]
    if missing_cols:
        print(f"ERRO: Colunas obrigatórias do HubSpot não encontradas: {missing_cols}")
        print("   Verifique se o arquivo HubSpot está completo.")
        sys.exit(1)
    
    # Detectar coluna de ID do HubSpot
    col_id = find_col(df_hub, ['record_id', 'id_do_registro', 'deal_id', 'object_id', 'id', 'recordid'])

    # Renomear e padronizar colunas
    rename_map = {
        'data_de_criacao': 'Data',
        'etapa_do_negocio': 'Status_Principal',
        'unidade_desejada': 'Unidade',
        'valor_na_moeda_da_empresa': 'RVO',
        'data_de_fechamento': 'Data_Fechamento',
        'fonte_original_do_trafego': 'Fonte_Original_do_Trafego',
        'detalhamento_da_fonte_original_do_trafego_1': 'Detalhamento_fonte_original_1',
        'detalhamento_da_fonte_original_do_trafego_2': 'Detalhamento_fonte_original_2',
        'pipeline': 'Tipo'
    }
    
    if col_id:
        print(f"   Coluna de ID encontrada: {col_id} -> Lead_Key")
        rename_map[col_id] = 'Lead_Key'
    else:
        print("   AVISO: Coluna de ID do HubSpot NÃO encontrada. IDs serão gerados sequencialmente.")

    df_hub = df_hub.rename(columns=rename_map)
    
    # Garantir existência da coluna Lead_Key
    if 'Lead_Key' not in df_hub.columns:
        df_hub['Lead_Key'] = "GEN_" + df_hub.index.astype(str)
    
    # Normalizar Lead_Key para string
    df_hub['Lead_Key'] = df_hub['Lead_Key'].astype(str).str.replace(r'\.0$', '', regex=True)

    # Converter Data de Criação para datetime (mantendo a hora para o ID)
    df_hub['Data'] = pd.to_datetime(df_hub['Data'], errors='coerce')
    df_hub['Data_Fechamento'] = pd.to_datetime(df_hub['Data_Fechamento'], errors='coerce').dt.normalize()

    # --- DEDUPLICAÇÃO POR LEAD_KEY ---
    # O export do HubSpot pode ter múltiplas linhas por negócio (ex: um registro por contato
    # associado ao deal). Aqui deduplicamos mantendo a linha com Data_Fechamento preenchida;
    # se ambas têm (ou nenhuma tem) data, mantemos a última ocorrência.
    n_antes = len(df_hub)
    df_hub = (
        df_hub
        .sort_values('Data_Fechamento', na_position='first')  # NaT primeiro → keep='last' prefere data preenchida
        .drop_duplicates(subset=['Lead_Key'], keep='last')
    )
    n_apos = len(df_hub)
    if n_antes > n_apos:
        print(f"   DEDUPLICAÇÃO: {n_antes - n_apos} linhas removidas por Lead_Key duplicado "
              f"({n_antes} → {n_apos} registros únicos).")
        print("   (Possível causa: export HubSpot com múltiplos contatos por negócio)")
    else:
        print(f"   Lead_Key: sem duplicatas no CSV ({n_apos} registros).")

    # Limpar e mapear RVO
    df_hub['RVO'] = df_hub['RVO'].astype(str).str.replace(r'[^\d,]', '', regex=True).str.replace(',', '.', regex=False)
    df_hub['RVO'] = pd.to_numeric(df_hub['RVO'], errors='coerce').fillna(0)
    
    # Limpar colunas de atribuição
    df_hub['Fonte_Original_do_Trafego_clean'] = df_hub['Fonte_Original_do_Trafego'].apply(clean_text)
    
    # --- ALTERAÇÃO: NÃO FILTRAR MAIS (Manter Orgânico + Pago) ---
    # Mantemos a variável df_hub_paid por compatibilidade, mas ela agora contém TODOS os leads
    df_hub_paid = df_hub.copy()
    print(f"   Total de Leads Carregados (Pago + Orgânico): {len(df_hub_paid)}")
    
    # Listas de palavras-chave para classificação (baseado em CANAL_FILTRO)
    KEYWORDS_SOCIAL = ['social pago', 'paid social', 'facebook', 'instagram', 'linkedin']
    KEYWORDS_SEARCH = ['pesquisa paga', 'paid search', 'cpc', 'google ads']

    def classify_channel(text):
        t = str(text).lower()
        if any(k in t for k in KEYWORDS_SOCIAL): return 'Social Pago'
        if any(k in t for k in KEYWORDS_SEARCH): return 'Pesquisa Paga'
        return 'Orgânico/Outros'

    # Mapear origem principal
    df_hub_paid['Origem_Principal'] = df_hub_paid['Fonte_Original_do_Trafego_clean'].apply(classify_channel)
    
    # Detectar matrículas ANTES do mapeamento.
    # Os valores do CSV têm sufixo de pipeline ex: "MATRÍCULA CONCLUÍDA (Red Balloon - Unidades de Rua)"
    # clean_text remove acentos e caracteres especiais, mas mantém o texto do sufixo — usar str.contains.
    _status_clean = df_hub_paid['Status_Principal'].apply(clean_text)
    _is_matricula_raw = _status_clean.str.contains('matricula concluida', na=False)

    # Mapear status principal (para exibição e agrupamento).
    # Extrai só a parte antes do "(" para eliminar o sufixo de pipeline antes do .map().
    # Usa clean_text para normalizar tanto o valor quanto as chaves do mapa.
    _ETAPA_FUNIL_MAP_CLEAN = {clean_text(k): v for k, v in ETAPA_FUNIL_MAP.items()}
    _status_before_paren = df_hub_paid['Status_Principal'].astype(str).str.split('(').str[0].str.strip()
    _status_key = _status_before_paren.apply(clean_text)
    df_hub_paid['Status_Principal'] = _status_key.map(_ETAPA_FUNIL_MAP_CLEAN).fillna(_status_before_paren)

    # Flag de matrícula: status contém "matrícula concluída" E tem data de fechamento
    df_hub_paid['Matriculas'] = (
        _is_matricula_raw & df_hub_paid['Data_Fechamento'].notna()
    ).astype(int)
    
    # Calcular ciclo de captação baseado na Data_Fechamento
    df_hub_paid['Ciclo_Captacao'] = calcular_ciclo_captacao(df_hub_paid['Data_Fechamento'])
    
    # Criar colunas de matrícula por ciclo
    ciclos_unicos = sorted(df_hub_paid['Ciclo_Captacao'].unique())
    ciclos_validos = [c for c in ciclos_unicos if c != DEFAULT_NA_TEXT]
    matriculas_ciclo_cols = []
    for ciclo in ciclos_validos:
        col_name = f"Matriculas_{ciclo.replace('.', '_').replace(' ', '_')}"
        df_hub_paid[col_name] = (
            (df_hub_paid['Ciclo_Captacao'] == ciclo) & (df_hub_paid['Matriculas'] == 1)
        ).astype(int)
        matriculas_ciclo_cols.append(col_name)
    
    # --- 3.2. Carregar Investimentos (Meta e Google) ---
    
    # ... (Lógica de carregamento de Meta e Google omitida para brevidade, mas deve ser mantida) ...
    # Assumindo que df_invest_daily é o resultado do merge de Meta e Google
    
    # Lógica de carregamento de Meta e Google (simplificada para o script)
    # **ATENÇÃO: ESTA PARTE DEVE SER COMPLETA NO SEU SCRIPT REAL**
    try:
        # Usar abas YoY que têm dados agregados por dia
        df_meta_raw = pd.read_excel(META_REPORT_FILE, sheet_name="Meta_YoY")
        df_google_raw = pd.read_excel(GOOGLE_REPORT_FILE, sheet_name="Google_YoY")
        
        # Meta: colunas são 'Data' e 'Investimento'
        df_meta = df_meta_raw.copy()
        df_meta = df_meta.rename(columns={'Investimento': 'Investimento_Meta'})
        
        # Google: colunas são 'Data' e 'Investimento_Google'
        df_google = df_google_raw[['Data', 'Investimento_Google']].copy()
        
        # Normalizar datas
        df_meta['Data'] = pd.to_datetime(df_meta['Data']).dt.normalize()
        df_google['Data'] = pd.to_datetime(df_google['Data']).dt.normalize()
        
        df_invest_daily = pd.merge(df_meta[['Data', 'Investimento_Meta']], df_google[['Data', 'Investimento_Google']], on='Data', how='outer').fillna(0)
        df_invest_daily['Investimento_Total'] = df_invest_daily['Investimento_Meta'] + df_invest_daily['Investimento_Google']
        
    except Exception as e:
        print(f"ERRO ao carregar arquivos de investimento (Meta/Google). Certifique-se de que os scripts Meta e Google foram executados. Detalhe: {e}")
        sys.exit(1)
    
    # --- 3.3. Prorrateio de Investimento (CORRIGIDO) ---
    
    # Adicionar coluna Data_Dia (apenas a data) para o merge de prorrateio
    df_hub_paid['Data_Dia'] = df_hub_paid['Data'].dt.normalize()
    
    # Agrupar leads pagos por dia e canal para o prorrateio
    df_leads_paid_daily_agg = df_hub_paid.groupby(['Data_Dia', 'Origem_Principal']).size().reset_index(name='Leads_Dia_Canal')
    
    # CORREÇÃO: Merge separado por canal para evitar trazer investimentos de outros canais
    # 1. Leads de Meta (Social Pago)
    df_meta = df_leads_paid_daily_agg[df_leads_paid_daily_agg['Origem_Principal'] == 'Social Pago'].copy()
    df_meta_invest = df_invest_daily[['Data', 'Investimento_Meta']].copy()
    df_meta_invest['Data'] = pd.to_datetime(df_meta_invest['Data']).dt.normalize()
    df_meta = pd.merge(df_meta, df_meta_invest, left_on='Data_Dia', right_on='Data', how='left').fillna(0)
    
    # Calcular o total de leads de Meta por dia
    df_meta['Leads_Meta_Dia_Total'] = df_meta.groupby('Data_Dia')['Leads_Dia_Canal'].transform('sum')
    
    # Calcular o CPL de Meta por dia (Investimento Meta / Total de Leads Meta)
    df_meta['CPL_Meta'] = df_meta['Investimento_Meta'] / df_meta['Leads_Meta_Dia_Total']
    df_meta['CPL_Meta'] = df_meta['CPL_Meta'].fillna(0).replace([np.inf, -np.inf], 0)
    
    # 2. Leads de Google (Pesquisa Paga)
    df_google = df_leads_paid_daily_agg[df_leads_paid_daily_agg['Origem_Principal'] == 'Pesquisa Paga'].copy()
    df_google_invest = df_invest_daily[['Data', 'Investimento_Google']].copy()
    df_google_invest['Data'] = pd.to_datetime(df_google_invest['Data']).dt.normalize()
    df_google = pd.merge(df_google, df_google_invest, left_on='Data_Dia', right_on='Data', how='left').fillna(0)
    
    # Calcular o total de leads de Google por dia
    df_google['Leads_Google_Dia_Total'] = df_google.groupby('Data_Dia')['Leads_Dia_Canal'].transform('sum')
    
    # Calcular o CPL de Google por dia (Investimento Google / Total de Leads Google)
    df_google['CPL_Google'] = df_google['Investimento_Google'] / df_google['Leads_Google_Dia_Total']
    df_google['CPL_Google'] = df_google['CPL_Google'].fillna(0).replace([np.inf, -np.inf], 0)
    
    # 3. Consolidar CPLs
    df_cpl_meta = df_meta[['Data_Dia', 'Origem_Principal', 'CPL_Meta']].copy()
    df_cpl_google = df_google[['Data_Dia', 'Origem_Principal', 'CPL_Google']].copy()
    
    # Renomear CPLs para Midia_Paga e concatenar
    df_cpl_meta = df_cpl_meta.rename(columns={'CPL_Meta': 'Midia_Paga'})
    df_cpl_google = df_cpl_google.rename(columns={'CPL_Google': 'Midia_Paga'})
    
    df_cpl_consolidado = pd.concat([df_cpl_meta, df_cpl_google], ignore_index=True)
    
    # 4. Merge de volta com o dataframe granular de leads pagos
    # Merge para trazer o CPL (Midia_Paga) calculado por dia/canal
    # O merge adiciona a coluna 'Midia_Paga' do df_cpl_consolidado ao df_hub_paid
    df_hub_paid = pd.merge(df_hub_paid, df_cpl_consolidado, on=['Data_Dia', 'Origem_Principal'], how='left')
    
    # A coluna 'Midia_Paga' agora contém o CPL correto. Aplicar fillna(0).
    df_hub_paid['Midia_Paga'] = df_hub_paid['Midia_Paga'].fillna(0)
    
    # --- CORREÇÃO: Placeholders para dias COM investimento mas SEM leads ---
    # Identificar dias com investimento por canal
    dias_invest_meta = set(df_invest_daily[df_invest_daily['Investimento_Meta'] > 0]['Data'].dt.normalize())
    dias_invest_google = set(df_invest_daily[df_invest_daily['Investimento_Google'] > 0]['Data'].dt.normalize())
    
    # Identificar dias com leads por canal
    dias_leads_meta = set(df_hub_paid[df_hub_paid['Origem_Principal'] == 'Social Pago']['Data_Dia'].unique())
    dias_leads_google = set(df_hub_paid[df_hub_paid['Origem_Principal'] == 'Pesquisa Paga']['Data_Dia'].unique())
    
    # Dias sem leads mas com investimento
    dias_meta_sem_leads = dias_invest_meta - dias_leads_meta
    dias_google_sem_leads = dias_invest_google - dias_leads_google
    
    placeholders = []
    for dia in dias_meta_sem_leads:
        invest = df_invest_daily[df_invest_daily['Data'].dt.normalize() == dia]['Investimento_Meta'].sum()
        if invest > 0:
            placeholders.append({
                'Lead_Key': f"SEM_LEAD_META_{dia.strftime('%Y%m%d')}",
                'Data': pd.Timestamp(dia),
                'Data_Dia': pd.Timestamp(dia),
                'Data_Fechamento': pd.NaT,
                'Unidade': 'Sem Lead',
                'Tipo': 'Sem Lead',
                'Status_Principal': 'Sem Lead',
                'Origem_Principal': 'Social Pago',
                'Fonte_Original_do_Trafego': 'Social pago',
                'Detalhamento_fonte_original_1': 'Sem Lead',
                'Detalhamento_fonte_original_2': 'Sem Lead',
                'RVO': 0,
                'Matriculas': 0,
                'Midia_Paga': invest,
                'Ciclo_Captacao': calcular_ciclo_captacao(pd.Series([pd.Timestamp(dia)])).iloc[0]
            })
    
    for dia in dias_google_sem_leads:
        invest = df_invest_daily[df_invest_daily['Data'].dt.normalize() == dia]['Investimento_Google'].sum()
        if invest > 0:
            placeholders.append({
                'Lead_Key': f"SEM_LEAD_GOOGLE_{dia.strftime('%Y%m%d')}",
                'Data': pd.Timestamp(dia),
                'Data_Dia': pd.Timestamp(dia),
                'Data_Fechamento': pd.NaT,
                'Unidade': 'Sem Lead',
                'Tipo': 'Sem Lead',
                'Status_Principal': 'Sem Lead',
                'Origem_Principal': 'Pesquisa Paga',
                'Fonte_Original_do_Trafego': 'Pesquisa paga',
                'Detalhamento_fonte_original_1': 'Sem Lead',
                'Detalhamento_fonte_original_2': 'Sem Lead',
                'RVO': 0,
                'Matriculas': 0,
                'Midia_Paga': invest,
                'Ciclo_Captacao': calcular_ciclo_captacao(pd.Series([pd.Timestamp(dia)])).iloc[0]
            })
    
    if placeholders:
        df_placeholders = pd.DataFrame(placeholders)
        # Adicionar colunas de matrícula por ciclo (valor 0)
        for col in matriculas_ciclo_cols:
            df_placeholders[col] = 0
        df_hub_paid = pd.concat([df_hub_paid, df_placeholders], ignore_index=True)
        print(f"    {len(placeholders)} dias sem leads adicionados (Meta: {len(dias_meta_sem_leads)}, Google: {len(dias_google_sem_leads)})")
    
    # Remover coluna auxiliar
    df_hub_paid = df_hub_paid.drop(columns=['Data_Dia'])
    
    # --- 3.4. Preparar DataFrame Final (Visao_Granular_Final) ---
    
    df_final_granular = df_hub_paid.copy()
    
    # Adicionar coluna Total_Negocios (1 para leads, 0 para placeholders)
    df_final_granular['Total_Negocios'] = 1
    df_final_granular.loc[df_final_granular['Unidade'] == 'Sem Lead', 'Total_Negocios'] = 0
    
    # Renomear colunas para o formato final (para o ID)
    df_final_granular = df_final_granular.rename(columns={
        'Detalhamento_fonte_original_1': 'Campanha_Raw_1',
        'Detalhamento_fonte_original_2': 'Campanha_Raw_2'
    })
    
    # --- LÓGICA DE CORREÇÃO DA COLUNA 'Campanha' (VERSÃO 2.0 - MAIS ROBUSTA) ---
    # O objetivo é garantir que a coluna 'Campanha' contenha o detalhamento mais específico.
    # Se Campanha_Raw_1 for um termo genérico, usamos Campanha_Raw_2 como Campanha.
    
    # 1. Definir termos genéricos (usando clean_text para comparação)
    # Inclui termos de Meta e Google que são genéricos (ex: facebook, instagram, cpc, Auto-tagged PPC)
    termos_genericos_clean = ['facebook', 'instagram', 'cpc', 'auto-tagged ppc', 'google', 'google ads']
    
    # 2. Criar a máscara de correção
    # Verifica se Campanha_Raw_1 limpo está na lista de termos genéricos
    campanha_raw_1_clean = df_final_granular['Campanha_Raw_1'].apply(clean_text)
    is_generico_mask = campanha_raw_1_clean.isin(termos_genericos_clean)
    
    # 3. Aplicar a lógica de correção: se Campanha_Raw_1 for genérico, usa Campanha_Raw_2
    needs_correction = is_generico_mask & df_final_granular['Campanha_Raw_2'].notna()
    
    # 4. Criar Campanha e Termo
    # Por padrão, Campanha = Campanha_Raw_1 e Termo = Campanha_Raw_2
    df_final_granular['Campanha'] = df_final_granular['Campanha_Raw_1']
    df_final_granular['Termo'] = df_final_granular['Campanha_Raw_2']
    
    # Onde precisa de correção:
    # Campanha recebe Campanha_Raw_2
    df_final_granular.loc[needs_correction, 'Campanha'] = df_final_granular.loc[needs_correction, 'Campanha_Raw_2']
    # Termo recebe Campanha_Raw_1
    df_final_granular.loc[needs_correction, 'Termo'] = df_final_granular.loc[needs_correction, 'Campanha_Raw_1']
    
    # 5. Remover colunas auxiliares
    df_final_granular = df_final_granular.drop(columns=['Campanha_Raw_1', 'Campanha_Raw_2'])
    
    # Renomear colunas para o formato final (para o ID)
    # As colunas 'Campanha' e 'Termo' já estão criadas e corrigidas acima,
    # então este bloco de renomeação não é mais necessário.
    # O ID usa as colunas 'Campanha' e 'Termo' que acabamos de criar/corrigir.
    pass
    
    # Garantir que as colunas de matrícula por ciclo sejam int (para evitar o erro de tipo misturado no groupby)
    for col in matriculas_ciclo_cols:
        df_final_granular[col] = pd.to_numeric(df_final_granular[col], errors='coerce').fillna(0).astype(int)
    
    # Validar unicidade
    total_leads = len(df_final_granular)
    unique_natural_ids = df_final_granular['Lead_Key'].nunique()
    
    if unique_natural_ids == total_leads:
        print("   Lead_Key (HubSpot ID): 100% único e validado!")
    else:
        print(f"   AVISO: {total_leads - unique_natural_ids} Lead_Keys duplicados encontrados.")
        # Não abortar, pois IDs duplicados podem existir no HubSpot se houver reimportação ou erro na fonte, mas é bom avisar.
    
    # --- 3.6. Finalização e Agregação ---
    
    # Mapeamento Final das Contas
    def get_conta_final(origem):
        if origem == 'Pesquisa Paga': return GOOGLE_ACCOUNT_LABEL
        if origem == 'Social Pago': return META_ACCOUNT_OTHER_LABEL
        return 'Orgânico / Outros'

    df_final_granular['Nome_Conta_Final'] = df_final_granular['Origem_Principal'].apply(get_conta_final)
    
    # Mapeamento da Área de Gestão RVO
    df_final_granular['Area_Gestao_RVO'] = AREA_GESTAO_DEFAULT
    
    # Reorganização e Limpeza (Garantir que colunas de texto não fiquem NaN)
    text_cols = ['Origem_Principal', 'Fonte_Original_do_Trafego', 'Campanha', 'Termo', 'Unidade', 'Tipo', 'Status_Principal', 'Nome_Conta_Final', 'Area_Gestao_RVO', 'Ciclo_Captacao']
    for col in text_cols:
        if col in df_final_granular.columns:
            df_final_granular[col] = df_final_granular[col].astype(str).fillna(DEFAULT_NA_TEXT)
            df_final_granular[col] = df_final_granular[col].replace('', DEFAULT_NA_TEXT)
            df_final_granular[col] = df_final_granular[col].replace('0', DEFAULT_NA_TEXT)
            
    # Garantir que a coluna 'Matriculas' seja int (para evitar o erro de tipo misturado no groupby)
    df_final_granular['Matriculas'] = pd.to_numeric(df_final_granular['Matriculas'], errors='coerce').fillna(0).astype(int)
    
    # NÃO aplicar fillna aqui — Data_Fechamento deve permanecer como datetime para as
    # agregações de matrículas e groupby de leads que vêm a seguir.
    # O fillna('Em Andamento') será aplicado apenas para exibição, após todas as agregações.

    # Colunas para a Visao_Granular_Final
    cols_granular = [
        'Lead_Key', 'Data', 'Data_Fechamento', 'Ciclo_Captacao', 'Unidade', 'Tipo', 'Total_Negocios', 
        'Midia_Paga', 'RVO', 'Matriculas', 'Status_Principal', 'Origem_Principal', 
        'Campanha', 'Termo', 'Fonte_Original_do_Trafego', 'Nome_Conta_Final', 'Area_Gestao_RVO'
    ] + matriculas_ciclo_cols
    
    cols_granular_exist = [c for c in cols_granular if c in df_final_granular.columns]
    df_final_granular = df_final_granular[cols_granular_exist]
    
    # --- 3.7. Geração das Abas Finais ---
    
    # 1. Agregação de Custo e Leads (ANCORADA NA CRIAÇÃO)
    # Colunas de agrupamento: todas as colunas não-métricas e não-ID
    # Colunas de agrupamento: todas as colunas não-métricas e não-ID
    # Excluir explicitamente as colunas de métrica e o ID
    metric_cols_to_exclude = ['Matriculas', 'Data_Fechamento', 'Midia_Paga', 'RVO', 'Total_Negocios', 'Lead_Key'] + matriculas_ciclo_cols
    group_cols_leads = [c for c in cols_granular_exist if c not in metric_cols_to_exclude]
    
    # Forçar o tipo de dado para string nas colunas de agrupamento para evitar o erro de tipo misturado
    for col in group_cols_leads:
        if df_final_granular[col].dtype != 'datetime64[ns]':
            df_final_granular[col] = df_final_granular[col].astype(str)
    
    # CORREÇÃO: dropna=False garante que leads com Data de criação nula (NaT) não sejam
    # descartados silenciosamente pelo groupby — evita a perda dos ~564 leads faltantes.
    df_agregado_leads = df_final_granular.groupby(group_cols_leads, dropna=False).agg(
        Investimento=('Midia_Paga', 'sum'),
        Volume_Total_Negocios=('Total_Negocios', 'sum'),
        RVO_Total=('RVO', 'sum')
    ).reset_index()
    
    df_agregado_leads = df_agregado_leads.rename(columns={'Data': 'Data_Criacao'})
    
    # 2. Agregação de Matrículas por DATA DE FECHAMENTO — SOMENTE MÍDIA PAGA
    # Regras:
    #   - Matriculas=1 e Data_Fechamento válida
    #   - Apenas canais pagos (Social Pago + Pesquisa Paga); orgânico não entra no relatório de mídia paga
    #   - Data_Fechamento <= hoje (evita datas futuras inválidas de vazar para o relatório)
    CANAIS_PAGOS = ['Social Pago', 'Pesquisa Paga']
    HOJE = pd.Timestamp.today().normalize()

    df_matriculas_validas = df_final_granular[
        (df_final_granular['Matriculas'] == 1) &
        (df_final_granular['Data_Fechamento'].notna()) &
        (df_final_granular['Origem_Principal'].isin(CANAIS_PAGOS)) &
        (pd.to_datetime(df_final_granular['Data_Fechamento'], errors='coerce') <= HOJE)
    ].copy()
    
    # Colunas de agrupamento para matrículas (Data_Fechamento é a âncora)
    # Colunas de agrupamento para matrículas (Data_Fechamento é a âncora)
    # Excluir 'Data' (Data de Criação) e incluir 'Data_Fechamento'
    group_cols_matriculas = [c for c in group_cols_leads if c != 'Data'] + ['Data_Fechamento']
    
    # Forçar o tipo de dado para string nas colunas de agrupamento (exceto Data_Fechamento)
    for col in group_cols_matriculas:
        if col != 'Data_Fechamento': # <--- Mantém Data_Fechamento como datetime
            df_matriculas_validas[col] = df_matriculas_validas[col].astype(str)
            
    # Garantir que Data_Fechamento seja datetime (embora já devesse ser)
    df_matriculas_validas['Data_Fechamento'] = pd.to_datetime(df_matriculas_validas['Data_Fechamento'], errors='coerce')
    
    matriculas_agg_dict = {'Matriculas': 'sum'}
    for col in matriculas_ciclo_cols:
        matriculas_agg_dict[col] = 'sum'
        
    # CORREÇÃO: dropna=False para não perder matrículas com alguma dimensão nula no groupby.
    df_agregado_matriculas = df_matriculas_validas.groupby(group_cols_matriculas, dropna=False).agg(
        matriculas_agg_dict
    ).reset_index()

    df_agregado_matriculas = df_agregado_matriculas.rename(columns={'Data_Fechamento': 'Data_Fechamento'})

    # --- CRIAR ABA MATRICULA_DATA_FECHAMENTO ---
    # Granular completo de negócios na etapa Matrícula Concluída, ancorado em Data de Fechamento.
    # Deve ser criado ANTES do fillna('Em Andamento') para que Data_Fechamento ainda seja datetime.
    df_matricula_data_fechamento = df_final_granular[
        (df_final_granular['Matriculas'] == 1) &
        (df_final_granular['Data_Fechamento'].notna()) &
        (df_final_granular['Origem_Principal'].isin(CANAIS_PAGOS)) &
        (pd.to_datetime(df_final_granular['Data_Fechamento'], errors='coerce') <= HOJE)
    ].copy().sort_values('Data_Fechamento').reset_index(drop=True)

    mostrar_matriculas_concluidas(df_matricula_data_fechamento)

    # Aplicar fillna em Data_Fechamento SOMENTE AQUI, após todas as agregações baseadas em data.
    # Mover para cá evita que o fillna('Em Andamento') polua o groupby e o filtro .notna() acima.
    df_final_granular['Data_Fechamento'] = df_final_granular['Data_Fechamento'].fillna('Em Andamento')

    # --- RESUMO MENSAL DE MATRÍCULAS POR DATA DE FECHAMENTO ---
    try:
        CHECK_DIR = OUTPUT_DIR / "visão resumida"
        CHECK_DIR.mkdir(parents=True, exist_ok=True)

        df_mat_check: pd.DataFrame = df_matriculas_validas.copy()
        df_mat_check['Ano_Fech'] = pd.to_datetime(df_mat_check['Data_Fechamento'], errors='coerce').dt.year
        df_mat_check['Mes_Fech'] = pd.to_datetime(df_mat_check['Data_Fechamento'], errors='coerce').dt.month

        MESES = {1:'Jan',2:'Fev',3:'Mar',4:'Abr',5:'Mai',6:'Jun',
                 7:'Jul',8:'Ago',9:'Set',10:'Out',11:'Nov',12:'Dez'}

        print('\n' + '='*60)
        print('RESUMO MENSAL DE MATRÍCULAS — MÍDIA PAGA (Data Fechamento)')
        print('='*60)
        print(f"{'Mês':<10} | {'Pesquisa Paga':>15} | {'Social Pago':>13} | {'TOTAL':>7}")
        print('-'*60)

        total_geral = 0
        anos = sorted(df_mat_check['Ano_Fech'].dropna().unique().astype(int))
        for ano in anos:  # type: ignore[assignment]
            df_ano: pd.DataFrame = df_mat_check[df_mat_check['Ano_Fech'] == ano]
            total_ano = 0
            for mes in range(1, 13):
                df_mes: pd.DataFrame = df_ano[df_ano['Mes_Fech'] == mes]
                if df_mes.empty:
                    continue
                por_canal = df_mes.groupby('Origem_Principal')['Matriculas'].sum()
                srch = int(por_canal.get('Pesquisa Paga', 0))
                soc  = int(por_canal.get('Social Pago', 0))
                tot  = srch + soc
                total_ano += tot
                label = f"{MESES[mes]}/{str(ano)[2:]}"
                print(f"{label:<10} | {srch:>15} | {soc:>13} | {tot:>7}")
            print(f"{'Total '+str(ano):<10} | {'':>15} | {'':>13} | {total_ano:>7}")
            print('-'*60)
            total_geral += total_ano

        print(f"{'TOTAL GERAL':<10} | {'':>15} | {'':>13} | {total_geral:>7}")
        print('='*60)

        # Exporta CSV completo de conferência
        OUT_MAT = CHECK_DIR / 'matriculas_mensal_fechamento.csv'
        df_mat_check.to_csv(OUT_MAT, index=False, encoding='utf-8-sig')
        print(f'\nCSV de conferência exportado: {OUT_MAT}')

    except Exception as e:
        print(f'ERRO no resumo de matrículas: {e}')
    
    # 3. Base Final Agregada para o Looker (Aba 1)
    df_agregado = df_agregado_leads.copy()
    df_agregado = df_agregado.rename(columns={
        'Origem_Principal': 'Canal',
        'Campanha': 'Campanha',
        'Termo': 'Termo',
        'Status_Principal': 'Etapas_de_Negocios',
        'Tipo': 'Pipeline',
        'Unidade': 'Unidade_Desejada',
        'Nome_Conta_Final': 'Conta_Agregada' 
    })
    
    # 4. Salvar Blend Final
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    BLEND_FILE_NAME = f"{BLEND_BASE_NAME}_{timestamp}.xlsx"
    BLEND_FILE_PATH = OUTPUT_DIR / BLEND_FILE_NAME
    
    print(f"\nSalvando blend final em: {BLEND_FILE_PATH}")
    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(BLEND_FILE_PATH, engine='openpyxl') as writer: 
            # Aba 1: Granular Geral
            df_final_granular.to_excel(writer, sheet_name="Visao_Granular_Final", index=False, float_format="%.2f")
            print("   Aba 1 ('Visao_Granular_Final') salva.")
            
            # Aba 2: Agregado Geral (Custo/Leads - Criação)
            df_agregado.to_excel(writer, sheet_name="Blend_Agregado_Dash", index=False, float_format="%.2f")
            print("   Aba 2 ('Blend_Agregado_Dash' - Custo/Leads) salva.")
            
            # Aba 3: Agregado Matrículas (ANCORADO NA DATA DE FECHAMENTO)
            df_agregado_matriculas.rename(columns={'Data_Fechamento': 'Data'}).to_excel(writer, sheet_name="Agregado_Matriculas_Fechamento", index=False, float_format="%.2f")
            print("   Aba 3 ('Agregado_Matriculas_Fechamento' - Matrículas) salva.")

            # Aba 4: Granular de Matrículas Concluídas por Data de Fechamento
            df_matricula_data_fechamento.to_excel(writer, sheet_name="Matricula_Data_Fechamento", index=False, float_format="%.2f")
            print("   Aba 4 ('Matricula_Data_Fechamento' - Matrículas Concluídas) salva.")

            # Aba 5: Resumo Mensal de Matrículas por Data de Fechamento (SOMENTE MÍDIA PAGA)
            df_resumo_mensal = df_matriculas_validas.copy()
            df_resumo_mensal['Ano'] = pd.to_datetime(df_resumo_mensal['Data_Fechamento'], errors='coerce').dt.year
            df_resumo_mensal['Mes'] = pd.to_datetime(df_resumo_mensal['Data_Fechamento'], errors='coerce').dt.month
            df_resumo_mensal = df_resumo_mensal.dropna(subset=['Ano', 'Mes'])
            df_resumo_mensal['Ano'] = df_resumo_mensal['Ano'].astype(int)
            df_resumo_mensal['Mes'] = df_resumo_mensal['Mes'].astype(int)
            
            # Agregar por Ano, Mês
            resumo_agrupado = df_resumo_mensal.groupby(['Ano', 'Mes']).agg({
                'Matriculas': 'sum'
            }).reset_index()
            
            # Criar coluna de período formatado
            MESES_NOME = {1:'Janeiro',2:'Fevereiro',3:'Março',4:'Abril',5:'Maio',6:'Junho',
                         7:'Julho',8:'Agosto',9:'Setembro',10:'Outubro',11:'Novembro',12:'Dezembro'}
            resumo_agrupado['Periodo'] = resumo_agrupado['Mes'].map(MESES_NOME) + ' ' + resumo_agrupado['Ano'].astype(str)
            
            # Reordenar colunas
            df_resumo_final = resumo_agrupado[['Periodo', 'Ano', 'Mes', 'Matriculas']].sort_values(['Ano', 'Mes']).reset_index(drop=True)
            
            df_resumo_final.to_excel(writer, sheet_name="Resumo_Mensal_Matriculas", index=False)
            print("   Aba 5 ('Resumo_Mensal_Matriculas' - Agregado Mensal Mídia Paga) salva.")

    except Exception as e:
        print(f"ERRO AO SALVAR O EXCEL: {e}")
        print("   Verifique se o arquivo não está aberto em outro programa.")
        sys.exit(1)

    # --- GERAÇÃO DO ARQUIVO RESUMIDO PARA DASHBOARD ---
    RESUMIDO_DIR = OUTPUT_DIR / "visão resumida"
    RESUMIDO_DIR.mkdir(parents=True, exist_ok=True)
    BLEND_RESUMIDO_PATH = RESUMIDO_DIR / "hubspot_resumido_Dash.xlsx"
    print(f"\nSalvando arquivo resumido em: {BLEND_RESUMIDO_PATH}")
    try:
        df_agregado_out: pd.DataFrame = df_agregado
        df_granular_out: pd.DataFrame = df_final_granular
        df_mat_export: pd.DataFrame = df_agregado_matriculas.rename(columns={'Data_Fechamento': 'Data'})
        
        # Criar resumo mensal de matrículas
        df_resumo_mensal_out = df_matriculas_validas.copy()
        df_resumo_mensal_out['Ano'] = pd.to_datetime(df_resumo_mensal_out['Data_Fechamento'], errors='coerce').dt.year
        df_resumo_mensal_out['Mes'] = pd.to_datetime(df_resumo_mensal_out['Data_Fechamento'], errors='coerce').dt.month
        df_resumo_mensal_out = df_resumo_mensal_out.dropna(subset=['Ano', 'Mes'])
        df_resumo_mensal_out['Ano'] = df_resumo_mensal_out['Ano'].astype(int)
        df_resumo_mensal_out['Mes'] = df_resumo_mensal_out['Mes'].astype(int)
        
        resumo_agrupado_out = df_resumo_mensal_out.groupby(['Ano', 'Mes']).agg({
            'Matriculas': 'sum'
        }).reset_index()
        
        MESES_NOME = {1:'Janeiro',2:'Fevereiro',3:'Março',4:'Abril',5:'Maio',6:'Junho',
                     7:'Julho',8:'Agosto',9:'Setembro',10:'Outubro',11:'Novembro',12:'Dezembro'}
        resumo_agrupado_out['Periodo'] = resumo_agrupado_out['Mes'].map(MESES_NOME) + ' ' + resumo_agrupado_out['Ano'].astype(str)
        df_resumo_final_out = resumo_agrupado_out[['Periodo', 'Ano', 'Mes', 'Matriculas']].sort_values(['Ano', 'Mes']).reset_index(drop=True)
        
        with pd.ExcelWriter(BLEND_RESUMIDO_PATH, engine='openpyxl') as writer:  # type: ignore[arg-type]
            df_agregado_out.to_excel(writer, sheet_name="Blend_Agregado_Dash", index=False, float_format="%.2f")  # type: ignore[arg-type]
            df_granular_out.to_excel(writer, sheet_name="Visao_Granular_Final", index=False, float_format="%.2f")  # type: ignore[arg-type]
            df_mat_export.to_excel(writer, sheet_name="Agregado_Matriculas_Fechamento", index=False, float_format="%.2f")  # type: ignore[arg-type]
            # Aba granular completa: negócios Matrícula Concluída por Data de Fechamento (todos os canais)
            df_matricula_data_fechamento.to_excel(writer, sheet_name="Matricula_Data_Fechamento", index=False, float_format="%.2f")  # type: ignore[arg-type]
            # Aba resumo mensal de matrículas (mídia paga)
            df_resumo_final_out.to_excel(writer, sheet_name="Resumo_Mensal_Matriculas", index=False)  # type: ignore[arg-type]
        print("   Arquivo resumido salvo (5 abas: Blend_Agregado_Dash, Visao_Granular_Final, Agregado_Matriculas_Fechamento, Matricula_Data_Fechamento, Resumo_Mensal_Matriculas).")
    except Exception as e:
        print(f"ERRO AO SALVAR O EXCEL RESUMIDO: {e}")

    # --- RESUMO MENSAL DE INVESTIMENTO ---
    print("\n" + "="*70)
    print("RESUMO MENSAL DE INVESTIMENTO POR CANAL")
    print("="*70)
    
    df_final_granular['Ano'] = pd.to_datetime(df_final_granular['Data']).dt.year
    df_final_granular['Mes'] = pd.to_datetime(df_final_granular['Data']).dt.month
    
    anos_disponiveis = sorted(df_final_granular['Ano'].dropna().unique().astype(int))
    for ano in anos_disponiveis:
        df_ano = df_final_granular[df_final_granular['Ano'] == ano]
        if len(df_ano) == 0:
            continue
            
        print(f"\nANO {ano}")
        print("-"*70)
        print(f"{'Mês':<6} | {'Meta (Social Pago)':>20} | {'Google (Pesquisa Paga)':>22} | {'Total':>15}")
        print("-"*70)
        
        total_meta_ano = 0
        total_google_ano = 0
        
        for mes in range(1, 13):
            df_mes = df_ano[df_ano['Mes'] == mes]
            if len(df_mes) == 0:
                continue
            
            meta = df_mes[df_mes['Origem_Principal'] == 'Social Pago']['Midia_Paga'].sum()
            google = df_mes[df_mes['Origem_Principal'] == 'Pesquisa Paga']['Midia_Paga'].sum()
            total = meta + google
            
            total_meta_ano += meta
            total_google_ano += google
            
            print(f"{mes:>4}   | R$ {meta:>17,.2f} | R$ {google:>19,.2f} | R$ {total:>12,.2f}")
        
        print("-"*70)
        print(f"{'TOTAL':>6} | R$ {total_meta_ano:>17,.2f} | R$ {total_google_ano:>19,.2f} | R$ {total_meta_ano + total_google_ano:>12,.2f}")
    
    print("\n" + "="*70)
    
    # --- RESUMO DO FUNIL POR ANO (SAFRA) ---
    print("\n" + "="*100)
    print("RESUMO DO FUNIL POR ANO (SAFRA/CRIAÇÃO)")
    print("="*100)

    # Filtrar NaNs no Ano para visualização limpa e converter para int
    df_view = df_final_granular.dropna(subset=['Ano']).copy()
    df_view['Ano'] = df_view['Ano'].astype(int)
    
    resumo_safra = pd.crosstab(
        df_view['Status_Principal'], 
        df_view['Ano'], 
        margins=True, 
        margins_name='Total Geral'
    )
    
    pd.set_option('display.max_rows', None)
    pd.set_option('display.width', 1000)
    print(resumo_safra)
    print("="*100)

    print("\nProcesso de blend concluído com sucesso!")
    print(f"   Arquivo gerado: {BLEND_FILE_PATH.name}")

# Bloco de execução principal
try:
    main()
except Exception as e:
    print(f"\nERRO INESPERADO NA EXECUÇÃO: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)