#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
analise_performance_hubspot_CORRIGIDO_CANAL.py

✅ CORREÇÕES IMPLEMENTADAS:
    - Prorrateio por CANAL (Investimento Canal / Leads Canal do Dia).
    - Evita distorção de custos entre Meta e Google.
    - Resolve o problema de investimento zerado (ignora erros de nome de campanha).
    - Mantém estrutura idêntica ao relatório atual (IDs, colunas, abas).
"""

import pandas as pd
import numpy as np
import unicodedata
import re
import sys
import hashlib
from pathlib import Path
from datetime import datetime

print("="*80)
print("🚀 Iniciando Script de BLEND - PRORRATEIO POR CANAL (CORRIGIDO)")
print("="*80)

# --- 1. CONFIGURAÇÕES ---

try:
    BASE_DIR = Path(__file__).resolve().parent.parent
except NameError:
    BASE_DIR = Path.cwd()

DATA_DIR_HUBSPOT = BASE_DIR / "data" 
DATA_DIR_INVESTIMENTO = BASE_DIR / "outputs" 
OUTPUT_DIR = BASE_DIR / "output" 

OUTPUT_DIR.mkdir(exist_ok=True)

# Arquivos de entrada
HUBSPOT_FILE = DATA_DIR_HUBSPOT / "hubspot_dataset.csv" 
META_REPORT_FILE = DATA_DIR_INVESTIMENTO / "meta_dataset_dashboard.xlsx"
META_SHEET_NAME = "Meta_Completo" 
GOOGLE_REPORT_FILE = DATA_DIR_INVESTIMENTO / "google_dashboard.xlsx"
GOOGLE_SHEET_NAME = "Google_Completo" 

# Nome do arquivo de saída
BLEND_BASE_NAME = "dataset_geral_melhorado"

# Labels
META_ACCOUNT_OTHER_LABEL = "Red Balloon - Contas Meta"
GOOGLE_ACCOUNT_LABEL = "Google Ads"
AREA_GESTAO_DEFAULT = "Gestão Antiga" 

# Mapeamento de Canais (HubSpot -> Categoria de Investimento)
CANAL_MAP_FINAL = {
    'social pago': 'Social Pago',
    'facebook': 'Social Pago',
    'instagram': 'Social Pago',
    'linkedin': 'Social Pago',
    'paid social': 'Social Pago',
    
    'pesquisa paga': 'Pesquisa Paga',
    'cpc': 'Pesquisa Paga',
    'google': 'Pesquisa Paga',
    'google ads': 'Pesquisa Paga'
}

# Mapeamento de Status
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
DEFAULT_NA_TEXT = "Não Mapeado"


# --- 2. FUNÇÕES UTILITÁRIAS ---

def read_any(path: Path, sheet_name=0, skiprows=0) -> pd.DataFrame:
    print(f"    📂 Lendo: {path.name}")
    if not path.exists():
        print(f"    ❌ Arquivo não encontrado: {path}")
        return pd.DataFrame()
    
    try:
        if path.suffix in (".xls", ".xlsx"):
            return pd.read_excel(path, sheet_name=sheet_name, skiprows=skiprows)
        elif path.suffix == ".csv":
            for encoding in ['utf-8', 'latin1', 'cp1252']:
                try:
                    return pd.read_csv(path, engine="python", sep=None, encoding=encoding, skiprows=skiprows)
                except: continue
            return pd.read_csv(path, skiprows=skiprows)
    except Exception as e:
        print(f"    ❌ Erro ao ler: {e}")
        return pd.DataFrame()
    return pd.DataFrame()

def clean_cols(df):
    """Normaliza nomes de colunas"""
    cols = []
    for c in df.columns:
        c_str = str(c).strip().lower()
        c_norm = unicodedata.normalize('NFKD', c_str).encode('ascii', 'ignore').decode('utf-8')
        c_clean = re.sub(r'[^a-z0-9_]+', '_', c_norm)
        cols.append(c_clean)
    df.columns = cols
    return df

def clean_text(text):
    """Limpa strings para merge"""
    try:
        text_str = str(text).strip().lower()
        text_norm = unicodedata.normalize('NFKD', text_str).encode('ascii', 'ignore').decode('utf-8')
        return re.sub(r'[^a-z0-9 ]+', '', text_norm).strip()
    except: return ""

def extract_status_base(status_full):
    """Remove pipeline do nome do status"""
    try:
        if pd.isna(status_full): return DEFAULT_NA_TEXT
        return str(status_full).split('(')[0].strip().upper()
    except: return DEFAULT_NA_TEXT

def find_col(df, keywords):
    """Encontra coluna por palavra-chave"""
    for col in df.columns:
        c_low = str(col).lower()
        for k in keywords:
            if k.lower() in c_low: return col
    return None

def calcular_ciclo_captacao(date_series):
    """Calcula YY.1 Alta ou YY.2 Baixa"""
    def get_ciclo(dt):
        if pd.isna(dt): return DEFAULT_NA_TEXT
        try:
            mes, ano = dt.month, dt.year
            if mes >= 10: return f"{str(ano + 1)[2:]}.1 Alta"
            elif mes <= 3: return f"{str(ano)[2:]}.1 Alta"
            else: return f"{str(ano)[2:]}.2 Baixa"
        except: return DEFAULT_NA_TEXT
    return date_series.apply(get_ciclo)

def generate_unique_id(df):
    """Gera IDs únicos mantendo a lógica original"""
    print("    🔑 Gerando IDs únicos...")
    
    # Chave de agrupamento para sequenciamento
    df['Chave_ID'] = df['Data'].dt.strftime('%Y%m%d') + '_' + \
                     df['Unidade'].astype(str) + '_' + \
                     df['Origem_Principal'].astype(str)
    
    df = df.sort_values(by=['Chave_ID', 'RVO'], ascending=[True, False])
    df['Sequencia_Desempate'] = df.groupby('Chave_ID').cumcount() + 1
    
    def create_ids(row):
        # Preparação dos componentes do ID
        data_str = row['Data'].strftime('%Y%m%d')
        
        # Limpeza rigorosa para o ID (apenas letras maiúsculas)
        unid_clean = re.sub(r'[^a-zA-Z]', '', unicodedata.normalize('NFKD', str(row['Unidade'])).encode('ascii', 'ignore').decode('utf-8')).upper()
        unid_str = unid_clean[:10].ljust(10, 'X')
        
        orig_clean = re.sub(r'[^a-zA-Z]', '', unicodedata.normalize('NFKD', str(row['Origem_Principal'])).encode('ascii', 'ignore').decode('utf-8')).upper()
        orig_str = orig_clean[:5].ljust(5, 'X')
        
        seq_str = str(row['Sequencia_Desempate']).zfill(3)
        
        # Hash curto para unicidade
        stable_key = f"{row['Chave_ID']}_{seq_str}"
        hash_val = hashlib.sha1(stable_key.encode()).hexdigest()[:4].upper()
        
        id_completo = f"{data_str}_{unid_str}_{orig_str}_{seq_str}_{hash_val}"
        
        # Lead Key (formato curto: YY + Hash + Seq)
        id_curto = f"{data_str[2:4]}{hash_val}{seq_str}"
        
        return pd.Series([id_completo, id_curto])

    df[['ID_Negocio_Completo', 'Lead_Key']] = df.apply(create_ids, axis=1)
    return df.drop(columns=['Chave_ID', 'Sequencia_Desempate'])


# --- 3. LÓGICA PRINCIPAL ---

def main():
    
    # =========================================================================
    # 1. CARREGAR E TRATAR HUBSPOT
    # =========================================================================
    print("\n📥 Processando Base HubSpot...")
    df_hub_raw = read_any(HUBSPOT_FILE)
    if df_hub_raw.empty:
        print("❌ ERRO CRÍTICO: Base HubSpot vazia.")
        sys.exit(1)
        
    df_hub = clean_cols(df_hub_raw)
    
    # Identificar colunas chave
    col_data = find_col(df_hub, ['createdate', 'data_de_criacao', 'create_date', 'data'])
    col_status = find_col(df_hub, ['etapa_do_negocio', 'dealstage', 'status'])
    col_fonte = find_col(df_hub, ['fonte_original_do_trafego', 'original_source'])
    col_unidade = find_col(df_hub, ['unidade_desejada', 'unidade'])
    col_rvo = find_col(df_hub, ['valor', 'amount', 'rvo'])
    col_close = find_col(df_hub, ['closedate', 'data_de_fechamento'])
    
    if not col_data:
        print("❌ ERRO: Coluna de Data não encontrada no HubSpot.")
        sys.exit(1)

    # Padronização de Dados
    df_hub['Data'] = pd.to_datetime(df_hub[col_data], errors='coerce').dt.normalize()
    df_hub['Data_Fechamento'] = pd.to_datetime(df_hub[col_close], errors='coerce').dt.normalize() if col_close else pd.NaT
    
    # Fonte e Canal
    df_hub['Fonte_Clean'] = df_hub[col_fonte].astype(str).apply(clean_text) if col_fonte else ""
    df_hub['Origem_Principal'] = df_hub['Fonte_Clean'].map(CANAL_MAP_FINAL).fillna(DEFAULT_NA_TEXT)
    
    # ⚠️ FILTRO: Trabalhar apenas com leads de Mídia Paga para o relatório
    df_paid = df_hub[df_hub['Origem_Principal'].isin(CANAL_MAP_FINAL.values())].copy()
    print(f"    ✅ Leads de Mídia Paga filtrados: {len(df_paid)}")
    
    # Tratamento de Status
    df_paid['Status_Original'] = df_paid[col_status].fillna(DEFAULT_NA_TEXT) if col_status else DEFAULT_NA_TEXT
    df_paid['Status_Base'] = df_paid['Status_Original'].apply(extract_status_base)
    df_paid['Status_Principal'] = df_paid['Status_Base'].map(ETAPA_FUNIL_MAP).fillna(DEFAULT_NA_TEXT)
    
    # Outros Campos
    df_paid['Unidade'] = df_paid[col_unidade].fillna(DEFAULT_NA_TEXT) if col_unidade else DEFAULT_NA_TEXT
    df_paid['RVO'] = pd.to_numeric(df_paid[col_rvo], errors='coerce').fillna(0) if col_rvo else 0
    
    # Detalhamentos (Campanha e Termo)
    c_det1 = find_col(df_hub, ['detalhamento_fonte_original_1', 'source_data_1'])
    c_det2 = find_col(df_hub, ['detalhamento_fonte_original_2', 'source_data_2'])
    df_paid['Detalhamento_fonte_original_1'] = df_paid[c_det1].fillna("-") if c_det1 else "-"
    df_paid['Detalhamento_fonte_original_2'] = df_paid[c_det2].fillna("-") if c_det2 else "-"
    
    # Pipeline / Tipo
    c_tipo = find_col(df_hub, ['pipeline', 'tipo'])
    df_paid['Tipo'] = df_paid[c_tipo].fillna("Pipeline Padrão") if c_tipo else "Pipeline Padrão"
    
    # Métricas Calculadas
    df_paid['Ciclo_Captacao'] = calcular_ciclo_captacao(df_paid['Data'])
    df_paid['Ciclo_Captacao_Fechamento'] = calcular_ciclo_captacao(df_paid['Data_Fechamento'])
    df_paid['Total_Negocios'] = 1
    df_paid['Matriculas'] = np.where(df_paid['Status_Principal'] == MATRICULA_NOME_FINAL, 1, 0)
    
    # Matrículas por Ciclo (Colunas Dinâmicas)
    for ciclo in df_paid['Ciclo_Captacao'].unique():
        if ciclo == DEFAULT_NA_TEXT: continue
        clean_ciclo = ciclo.replace('.', '_').replace(' ', '_')
        col_name = f'Matriculas_{clean_ciclo}'
        df_paid[col_name] = np.where(
            (df_paid['Matriculas']==1) & (df_paid['Ciclo_Captacao']==ciclo), 1, 0
        )

    # Campos de Conta/Gestão
    df_paid['Nome_Conta_Final'] = df_paid['Origem_Principal'].apply(
        lambda x: META_ACCOUNT_OTHER_LABEL if x == 'Social Pago' else GOOGLE_ACCOUNT_LABEL
    )
    df_paid['Area_Gestao_RVO'] = AREA_GESTAO_DEFAULT

    # =========================================================================
    # 2. CARREGAR E PREPARAR INVESTIMENTO (AGREGADO POR DIA)
    # =========================================================================
    print("\n💰 Processando Investimentos por Canal...")
    
    # --- META ADS (Social Pago) ---
    meta_daily = pd.DataFrame(columns=['Data', 'Investimento_Meta'])
    if META_REPORT_FILE.exists():
        df_meta = read_any(META_REPORT_FILE, sheet_name=META_SHEET_NAME)
        if not df_meta.empty:
            df_meta = clean_cols(df_meta)
            c_data = find_col(df_meta, ['data', 'date', 'day'])
            c_val = find_col(df_meta, ['investimento', 'spend', 'valor', 'amount'])
            
            if c_data and c_val:
                df_meta['Data'] = pd.to_datetime(df_meta[c_data], errors='coerce').dt.normalize()
                df_meta['Val'] = pd.to_numeric(df_meta[c_val], errors='coerce').fillna(0)
                # Agrupa por dia para ter o TOTAL DIÁRIO DO CANAL
                meta_daily = df_meta.groupby('Data')['Val'].sum().reset_index().rename(columns={'Val': 'Investimento_Meta'})
                print(f"    ✅ Meta Ads (Social Pago): R$ {meta_daily['Investimento_Meta'].sum():,.2f}")

    # --- GOOGLE ADS (Pesquisa Paga) ---
    google_daily = pd.DataFrame(columns=['Data', 'Investimento_Google'])
    if GOOGLE_REPORT_FILE.exists():
        df_google = read_any(GOOGLE_REPORT_FILE, sheet_name=GOOGLE_SHEET_NAME)
        if not df_google.empty:
            df_google = clean_cols(df_google)
            c_data = find_col(df_google, ['data', 'date', 'day'])
            c_val = find_col(df_google, ['investimento', 'cost', 'valor', 'spend'])
            
            if c_data and c_val:
                df_google['Data'] = pd.to_datetime(df_google[c_data], errors='coerce').dt.normalize()
                df_google['Val'] = pd.to_numeric(df_google[c_val], errors='coerce').fillna(0)
                # Agrupa por dia para ter o TOTAL DIÁRIO DO CANAL
                google_daily = df_google.groupby('Data')['Val'].sum().reset_index().rename(columns={'Val': 'Investimento_Google'})
                print(f"    ✅ Google Ads (Pesquisa Paga): R$ {google_daily['Investimento_Google'].sum():,.2f}")

    # =========================================================================
    # 3. CÁLCULO DO CPL DIÁRIO POR CANAL (PRORRATEIO CORRETO)
    # =========================================================================
    print("\n🧮 Calculando CPL Diário por Canal...")

    # Contar Leads por Dia e Canal NA BASE HUBSPOT
    # Isso nos dá o denominador da divisão (Investimento / Leads)
    leads_agg = df_paid.groupby(['Data', 'Origem_Principal']).size().reset_index(name='Qtd_Leads')

    # Separar contagens
    leads_social = leads_agg[leads_agg['Origem_Principal'] == 'Social Pago'].copy()
    leads_pesquisa = leads_agg[leads_agg['Origem_Principal'] == 'Pesquisa Paga'].copy()

    # Merge Investimento com Quantidade de Leads (Social)
    cpl_social = pd.merge(leads_social, meta_daily, on='Data', how='left').fillna(0)
    cpl_social['CPL_Calculado'] = np.where(
        cpl_social['Qtd_Leads'] > 0,
        cpl_social['Investimento_Meta'] / cpl_social['Qtd_Leads'],
        0
    )
    
    # Merge Investimento com Quantidade de Leads (Pesquisa)
    cpl_pesquisa = pd.merge(leads_pesquisa, google_daily, on='Data', how='left').fillna(0)
    cpl_pesquisa['CPL_Calculado'] = np.where(
        cpl_pesquisa['Qtd_Leads'] > 0,
        cpl_pesquisa['Investimento_Google'] / cpl_pesquisa['Qtd_Leads'],
        0
    )

    # Concatenar tabelas de CPL
    df_cpl_final = pd.concat([
        cpl_social[['Data', 'Origem_Principal', 'CPL_Calculado']],
        cpl_pesquisa[['Data', 'Origem_Principal', 'CPL_Calculado']]
    ])

    print("    ✅ CPLs calculados com sucesso (separado por canal).")

    # =========================================================================
    # 4. LEVAR CUSTO PARA A TABELA FINAL
    # =========================================================================
    
    # Merge na base analítica usando DATA + CANAL
    df_final = pd.merge(
        df_paid,
        df_cpl_final,
        on=['Data', 'Origem_Principal'],
        how='left'
    )
    
    # Preencher coluna Midia_Paga
    df_final['Midia_Paga'] = df_final['CPL_Calculado'].fillna(0)
    
    print(f"    💰 Investimento Total Atribuído na Base: R$ {df_final['Midia_Paga'].sum():,.2f}")

    # =========================================================================
    # 5. GERAÇÃO DE SAÍDAS
    # =========================================================================
    
    # Gerar IDs
    df_granular = generate_unique_id(df_final)
    
    # Definir colunas para exportação (Visão Granular)
    # Baseado no snippet do seu arquivo 'Visao_Granular_Final.csv'
    cols_base = [
        'ID_Negocio_Completo', 'Lead_Key', 'Data', 'Data_Fechamento', 'Ciclo_Captacao', 
        'Unidade', 'Tipo', 'Total_Negocios', 'Midia_Paga', 'RVO', 'Matriculas', 
        'Status_Principal', 'Origem_Principal', 'Detalhamento_fonte_original_1', 
        'Detalhamento_fonte_original_2', 'Fonte_Original_do_Trafego'
    ]
    
    # Adiciona as colunas de Matrícula dinâmicas
    cols_mat = sorted([c for c in df_granular.columns if c.startswith('Matriculas_')])
    
    # Outras colunas que aparecem no seu csv de exemplo
    cols_extra = ['Nome_Conta_Final', 'Area_Gestao_RVO']
    
    cols_final = [c for c in (cols_base + cols_extra + cols_mat) if c in df_granular.columns]
    
    df_export_granular = df_granular[cols_final]

    # --- Agregação para Dashboard (Blend_Agregado_Dash) ---
    print("    📊 Preparando aba Blend_Agregado_Dash...")
    
    # Agrupar pelos campos dimensionais
    dims_agg = ['Data', 'Origem_Principal', 'Detalhamento_fonte_original_1', 
                'Detalhamento_fonte_original_2', 'Status_Principal', 'Tipo', 'Unidade']
    # Validar se existem
    dims_agg = [d for d in dims_agg if d in df_export_granular.columns]
    
    agg_rules = {
        'Total_Negocios': 'sum',
        'Matriculas': 'sum',
        'Midia_Paga': 'sum',
        'RVO': 'sum'
    }
    # Adicionar regras para colunas de matricula
    for cm in cols_mat:
        agg_rules[cm] = 'sum'
        
    df_dash = df_export_granular.groupby(dims_agg, dropna=False).agg(agg_rules).reset_index()
    
    # Renomear colunas para bater com o seu CSV 'Blend_Agregado_Dash.csv'
    rename_map = {
        'Origem_Principal': 'Canal',
        'Detalhamento_fonte_original_1': 'Campanha',
        'Detalhamento_fonte_original_2': 'Termo',
        'Status_Principal': 'Etapas_de_Negocios',
        'Tipo': 'Pipeline',
        'Unidade': 'Unidade_Desejada',
        'Total_Negocios': 'Volume_Total_Negocios',
        'Midia_Paga': 'Investimento',
        'RVO': 'RVO_Total'
    }
    df_dash = df_dash.rename(columns=rename_map)
    
    # Reordenar colunas do Dash
    cols_dash_order = ['Data', 'Canal', 'Campanha', 'Termo', 'Etapas_de_Negocios', 'Pipeline', 
                       'Unidade_Desejada', 'Volume_Total_Negocios', 'Matriculas', 'Investimento', 'RVO_Total'] + cols_mat
    
    # Filtrar colunas existentes
    cols_dash_final = [c for c in cols_dash_order if c in df_dash.columns]
    df_dash = df_dash[cols_dash_final]

    # --- Agregação Matrículas Fechamento (Agregado_Matriculas_Fechamento) ---
    print("    🎓 Preparando aba Agregado_Matriculas_Fechamento...")
    
    df_mat_fech = df_granular[df_granular['Matriculas'] == 1].copy()
    
    if not df_mat_fech.empty:
        # Recalcular colunas de matricula com base no ciclo de FECHAMENTO para esta aba específica
        # (Conforme lógica usual desse tipo de relatório, mas mantendo simples se não pedido)
        # Se precisar manter as colunas de criação, mantemos. Se precisar recalcular:
        # Aqui vamos manter a agregação baseada nos dados já calculados, agrupando por Data FECHAMENTO
        
        dims_fech = ['Data_Fechamento', 'Ciclo_Captacao_Fechamento', 'Origem_Principal', 
                     'Detalhamento_fonte_original_1', 'Detalhamento_fonte_original_2', 'Tipo', 'Unidade']
        # Ajuste de nomes se necessário antes do groupby
        
        # Agregação
        df_agreg_fech = df_mat_fech.groupby([c for c in dims_fech if c in df_mat_fech.columns], dropna=False).agg(agg_rules).reset_index()
        
        # Renomear
        rename_fech = {
            'Origem_Principal': 'Canal',
            'Detalhamento_fonte_original_1': 'Campanha',
            'Detalhamento_fonte_original_2': 'Termo',
            'Tipo': 'Pipeline',
            'Unidade': 'Unidade_Desejada',
            'Total_Negocios': 'Volume_Matriculas', # Aqui muda o nome
            'Midia_Paga': 'Investimento',
            'RVO': 'RVO_Total',
            'Ciclo_Captacao_Fechamento': 'Ciclo_Captacao' # Renomeia para ficar igual ao header do CSV
        }
        df_agreg_fech = df_agreg_fech.rename(columns=rename_fech)
        
        # Ordenar colunas
        cols_fech_order = ['Data_Fechamento', 'Ciclo_Captacao', 'Canal', 'Campanha', 'Termo', 
                           'Pipeline', 'Unidade_Desejada', 'Volume_Matriculas', 'Matriculas', 'Investimento', 'RVO_Total'] + cols_mat
        
        cols_fech_final = [c for c in cols_fech_order if c in df_agreg_fech.columns]
        df_agreg_fech = df_agreg_fech[cols_fech_final]
    else:
        df_agreg_fech = pd.DataFrame()

    # =========================================================================
    # 6. SALVAR ARQUIVO
    # =========================================================================
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    OUT_FILE = OUTPUT_DIR / f"{BLEND_BASE_NAME}_{timestamp}.xlsx"
    
    print(f"\n💾 Salvando arquivo final em: {OUT_FILE}")
    
    try:
        with pd.ExcelWriter(OUT_FILE, engine='openpyxl') as writer:
            df_export_granular.to_excel(writer, sheet_name='Visao_Granular_Final', index=False)
            df_dash.to_excel(writer, sheet_name='Blend_Agregado_Dash', index=False)
            if not df_agreg_fech.empty:
                df_agreg_fech.to_excel(writer, sheet_name='Agregado_Matriculas_Fechamento', index=False)
            print("    ✅ Arquivo salvo com sucesso!")
            
        print("\n📊 RESUMO FINAL:")
        print(f"    - Total Negócios: {df_export_granular['Total_Negocios'].sum()}")
        print(f"    - Investimento Total: R$ {df_export_granular['Midia_Paga'].sum():,.2f}")
        
    except Exception as e:
        print(f"❌ Erro ao salvar: {e}")

if __name__ == "__main__":
    main()