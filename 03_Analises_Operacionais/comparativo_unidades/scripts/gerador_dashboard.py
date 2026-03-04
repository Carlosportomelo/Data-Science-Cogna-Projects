# -*- coding: utf-8 -*-
"""
SCRIPT GERADOR DE DASHBOARD HTML - RED BALLOON
Gera um dashboard interativo com base nos dados de leads.
"""

import pandas as pd
from pathlib import Path
from datetime import datetime
import json
import logging
import sys

# Configuração de Logs
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / 'outputs'

# ===== CONSTANTES (Mesmas do script principal para consistência) =====
UNIDADES_PROPRIAS = {
    'VILA LEOPOLDINA', 'MORUMBI', 'ITAIM BIBI', 'PACAEMBU', 'PINHEIROS', 
    'JARDINS', 'PERDIZES', 'SANTANA'
}

UNIDADES_FRANQUEADAS = {
    'TATUAPE', 'SALVADOR', 'GUARULHOS', 'JUNDIAI', 'ASA SUL', 'SAO BERNARDO DO CAMPO',
    'SANTO ANDRE', 'CURITIBA', 'SAO CAETANO DO SUL', 'POMPEIA', 'MOOCA', 'BOA VIAGEM',
    'GOIANIA', 'UBERLANDIA', 'TIJUCA', 'ALPHAVILLE', 'LOURDES', 'TAQUARAL',
    'RIBEIRAO PRETO', 'GRACAS', 'PIRACICABA', 'SAO JOSE DOS CAMPOS', 'JACAREPAGUA',
    'IPIRANGA', 'PAMPULHA', 'CASA FORTE', 'CAMPO BELO', 'VILA SAO FRANCISCO',
    'ALDEOTA', 'MOGI DAS CRUZES', 'AGUA FRIA', 'PARAISO', 'MOEMA', 'CUIABA',
    'NOVA CAMPINAS', 'VALINHOS', 'SAUDE', 'ATIBAIA', 'BUTANTA', 'SION',
    'SANTO AMARO', 'GRANJA VIANA', 'SOROCABA', 'LIMEIRA', 'LONDRINA', 'TUCURUVI',
    'AMERICANA', 'INDAIATUBA', 'BARRA JARDIM OCEANICO', 'BAURU', 'PORTO ALEGRE',
    'CHACARA KLABIN', 'ICARAI', 'NOVA LIMA', 'CAMPO GRANDE', 'BOSQUE DA BARRA',
    'VILA MARIANA', 'MANAUS', 'ANAPOLIS', 'BARRA BLUE SQUARE', 'IPANEMA',
    'ARACATUBA', 'REGIAO OCEANICA', 'SAO JOSE DO RIO PRETO', 'ITU'
}

def find_csv():
    """Encontra o arquivo CSV de base."""
    paths_to_check = [
        BASE_DIR / 'data',
        Path.cwd(),
        Path.cwd().parent,
        Path.home(),
        Path.home() / 'Downloads',
    ]
    unique_paths = sorted(list(set(paths_to_check)))
    for p in unique_paths:
        if p.exists():
            found_files = list(p.glob('base_leads*.csv'))
            if found_files:
                return sorted(found_files, key=lambda x: x.stat().st_mtime, reverse=True)[0]
    raise FileNotFoundError("Nenhum arquivo 'base_leads*.csv' encontrado.")

def process_data():
    """Lê e processa os dados para gerar os DataFrames necessários."""
    input_file = find_csv()
    logger.info(f"Lendo arquivo: {input_file}")
    
    df = None
    for encoding in ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252']:
        try:
            df = pd.read_csv(input_file, sep=',', encoding=encoding)
            break
        except:
            continue
            
    if df is None:
        raise ValueError("Não foi possível ler o arquivo CSV.")

    df.columns = df.columns.str.strip()

    # Cenário Total (Bruto)
    _etapa_raw = df['Etapa do negócio'].astype(str).str.strip().str.upper()
    raw_conv = _etapa_raw.str.contains('MATRÍCULA CONCLUÍDA', na=False).sum()
    raw_perd = _etapa_raw.str.contains('PERDIDO', na=False).sum()
    raw_total = len(df)

    # Filtros e Limpeza
    UNIDADES_ANALISADAS = UNIDADES_PROPRIAS.union(UNIDADES_FRANQUEADAS)
    df = df.dropna(subset=['Unidade Desejada']).copy()
    df['Unidade Desejada'] = df['Unidade Desejada'].astype(str).str.strip().str.upper()
    df = df[df['Unidade Desejada'].isin(UNIDADES_ANALISADAS)].copy()

    # Datas
    col_create = next((c for c in df.columns if c.lower() in ['create date', 'data de criação', 'data de criacao', 'createdate']), None)
    col_close = next((c for c in df.columns if c.lower() in ['close date', 'data de fechamento', 'closedate', 'closed date', 'close_date']), None)
    
    if col_create:
        df['dt_create'] = pd.to_datetime(df[col_create], dayfirst=True, errors='coerce')
        if col_close:
            df['dt_close'] = pd.to_datetime(df[col_close], dayfirst=True, errors='coerce')
            mask_error = (df['dt_create'].notna() & df['dt_close'].notna()) & (df['dt_close'] < df['dt_create'])
            df = df[~mask_error].copy()

    # Classificação e Conversão
    df['Tipo'] = df['Unidade Desejada'].apply(lambda x: 'Própria' if x in UNIDADES_PROPRIAS else 'Franqueada')
    df['Etapa_Norm'] = df['Etapa do negócio'].astype(str).str.strip().str.upper()
    df['Convertido'] = df['Etapa_Norm'].str.contains('MATRÍCULA CONCLUÍDA', na=False).astype(int)
    df['Perdido'] = df['Etapa_Norm'].str.contains('PERDIDO', na=False).astype(int)
    df['Número de atividades de vendas'] = pd.to_numeric(df['Número de atividades de vendas'], errors='coerce').fillna(0).astype(int)

    # Métricas Gerais (Válidas)
    total = len(df)
    conv = int(df['Convertido'].sum())
    perd = int(df['Perdido'].sum())

    # Análise Detalhada
    # Garante que a coluna de responsável existe
    if 'Proprietário do negócio' not in df.columns:
        df['Proprietário do negócio'] = 'N/A'

    analise = df.groupby(['Tipo', 'Unidade Desejada']).agg({
        'Record ID': 'count',
        'Convertido': 'sum',
        'Perdido': 'sum',
        'Número de atividades de vendas': 'mean',
        'Proprietário do negócio': lambda x: x.mode()[0] if len(x.mode()) > 0 else 'N/A'
    }).rename(columns={'Record ID': 'Total', 'Convertido': 'Matriculas', 'Perdido': 'Perdidos',
                       'Número de atividades de vendas': 'Ativ_Media_Total', 'Proprietário do negócio': 'Responsável'})
    
    ativ_media_mat = df[df['Convertido'] == 1].groupby(['Tipo', 'Unidade Desejada'])['Número de atividades de vendas'].mean()
    ativ_media_perd = df[df['Perdido'] == 1].groupby(['Tipo', 'Unidade Desejada'])['Número de atividades de vendas'].mean()
    analise = analise.join(ativ_media_mat.rename('Ativ_Media_Mat')).fillna({'Ativ_Media_Mat': 0})
    analise = analise.join(ativ_media_perd.rename('Ativ_Media_Perd')).fillna({'Ativ_Media_Perd': 0})
    
    analise['Taxa_Conv'] = (analise['Matriculas'] / analise['Total'] * 100).round(1)
    analise['Taxa_Perda'] = (analise['Perdidos'] / analise['Total'] * 100).round(1)
    analise['Leads_p_Matricula'] = (analise['Total'] / analise['Matriculas']).round(1)

    # Timeline
    timeline_data = pd.DataFrame()
    yearly_data = pd.DataFrame()
    if 'dt_create' in df.columns:
        # 1. Leads por Data de Criação (Total)
        df_leads = df[df['dt_create'].notna()].copy()
        df_leads['Periodo'] = df_leads['dt_create'].dt.to_period('M')
        leads_series = df_leads.groupby('Periodo')['Record ID'].count().rename('Total')
        
        # 2. Matrículas e Perdidos
        # Definir data de referência: dt_close se existir, senão dt_create (fallback para não zerar perdidos)
        if 'dt_close' in df.columns:
            df['dt_ref_event'] = df['dt_close'].fillna(df['dt_create'])
        else:
            df['dt_ref_event'] = df['dt_create']
            
        # Matrículas (Convertido=1)
        df_mat = df[(df['dt_ref_event'].notna()) & (df['Convertido'] == 1)].copy()
        df_mat['Periodo'] = df_mat['dt_ref_event'].dt.to_period('M')
        mat_series = df_mat.groupby('Periodo')['Record ID'].count().rename('Matriculas')
        
        # Perdidos (Perdido=1)
        df_perd = df[(df['dt_ref_event'].notna()) & (df['Perdido'] == 1)].copy()
        df_perd['Periodo'] = df_perd['dt_ref_event'].dt.to_period('M')
        perd_series = df_perd.groupby('Periodo')['Record ID'].count().rename('Perdidos')

        # Combinar as séries
        timeline_data = pd.concat([leads_series, mat_series, perd_series], axis=1).fillna(0)
        timeline_data.index = timeline_data.index.astype(str)
        
        # Yearly Data (Agregado do timeline para consistência)
        # Extrair ano do índice string 'YYYY-MM'
        timeline_data['Ano'] = timeline_data.index.str[:4].astype(int)
        yearly_data = timeline_data.groupby('Ano').sum()
        # Remove coluna auxiliar do timeline_data para não atrapalhar o gráfico
        timeline_data = timeline_data.drop(columns=['Ano'])

    # Canais
    analise_canais = pd.DataFrame()
    analise_canais_unidade = pd.DataFrame()
    col_fonte = None
    keywords = ['fonte original', 'original source', 'source', 'fonte', 'origem']
    for col in df.columns:
        if any(k in col.lower() for k in keywords):
            col_fonte = col
            break
    
    if col_fonte:
        df[col_fonte] = df[col_fonte].fillna('Não Identificado').astype(str).str.strip()
        # Geral
        analise_canais = df.groupby(col_fonte).agg({
            'Record ID': 'count',
            'Convertido': 'sum',
            'Perdido': 'sum'
        }).rename(columns={'Record ID': 'Total', 'Convertido': 'Matriculas', 'Perdido': 'Perdidos'}).sort_values('Total', ascending=False)
        
        # Por Unidade
        analise_canais_unidade = df.groupby(['Tipo', 'Unidade Desejada', col_fonte]).agg({
            'Record ID': 'count',
            'Convertido': 'sum',
            'Perdido': 'sum'
        }).rename(columns={'Record ID': 'Total', 'Convertido': 'Matriculas', 'Perdido': 'Perdidos'}).reset_index()
        analise_canais_unidade = analise_canais_unidade.rename(columns={col_fonte: 'Canal'})

    summary_data = [
        ("Total de Leads", total, raw_total, '#,##0'),
        ("Matrículas", conv, raw_conv, '#,##0'),
        ("Taxa de Conversão", conv / total if total > 0 else 0, raw_conv / raw_total if raw_total > 0 else 0, '0.0%'),
        ("Perdidos", perd, raw_perd, '#,##0'),
        ("Taxa de Perda", perd / total if total > 0 else 0, raw_perd / raw_total if raw_total > 0 else 0, '0.0%'),
        ("Em Pipeline", total - conv - perd, raw_total - raw_conv - raw_perd, '#,##0')
    ]

    return summary_data, analise, timeline_data, analise_canais, yearly_data, analise_canais_unidade

def generate_html(summary_data, analise, timeline_data, analise_canais, yearly_data, analise_canais_unidade):
    """Gera o arquivo HTML."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    html_file = OUTPUT_DIR / f"DASHBOARD_PERFORMANCE_{timestamp}.html"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Preparar dados JSON (Conversão explícita para tipos nativos do Python)
    # Formatar datas para "Jan 21" (PT-BR)
    def format_date_pt(date_str):
        try:
            dt = datetime.strptime(str(date_str), '%Y-%m')
            meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
            return f"{meses[dt.month - 1]} {dt.strftime('%y')}"
        except:
            return str(date_str)

    timeline_labels = [format_date_pt(x) for x in timeline_data.index] if not timeline_data.empty else []
    timeline_leads = [int(x) for x in timeline_data['Total'].tolist()] if not timeline_data.empty else []
    timeline_matriculas = [int(x) for x in timeline_data['Matriculas'].tolist()] if not timeline_data.empty else []
    timeline_perdidos = [int(x) for x in timeline_data['Perdidos'].tolist()] if not timeline_data.empty else []
    
    # Largura dinâmica para scroll horizontal (50px por mês ou mínimo de 800px)
    # Largura dinâmica para scroll horizontal (50px por mês ou mínimo de 800px)
    timeline_min_width = max(800, len(timeline_labels) * 50)
    
    if not analise_canais.empty:
        top_canais = analise_canais.head(10)
        canais_labels = top_canais.index.tolist()
        canais_matriculas = [int(x) for x in top_canais['Matriculas'].tolist()]
        canais_leads = [int(x) for x in top_canais['Total'].tolist()]
    else:
        canais_labels, canais_matriculas, canais_leads = [], [], []

    # Gerar HTML do Comparativo Anual
    yearly_cards_html = ""
    if not yearly_data.empty:
        years = yearly_data.index.tolist()
        for i, year in enumerate(years):
            curr = yearly_data.loc[year]
            
            # Deltas
            delta_leads_str = ""
            delta_mat_str = ""
            
            if i > 0:
                prev = yearly_data.loc[years[i-1]]
                
                # Leads Delta
                if prev['Total'] > 0:
                    d_leads = ((curr['Total'] - prev['Total']) / prev['Total']) * 100
                    color = "text-success" if d_leads >= 0 else "text-danger"
                    icon = "▲" if d_leads >= 0 else "▼"
                    delta_leads_str = f'<span class="{color} small ms-2" style="font-size: 0.6em;">{icon} {d_leads:.1f}%</span>'
                
                # Matriculas Delta
                if prev['Matriculas'] > 0:
                    d_mat = ((curr['Matriculas'] - prev['Matriculas']) / prev['Matriculas']) * 100
                    color = "text-success" if d_mat >= 0 else "text-danger"
                    icon = "▲" if d_mat >= 0 else "▼"
                    delta_mat_str = f'<span class="{color} small ms-2" style="font-size: 0.6em;">{icon} {d_mat:.1f}%</span>'

            conv_rate = (curr['Matriculas'] / curr['Total'] * 100) if curr['Total'] > 0 else 0
            
            yearly_cards_html += f"""
            <div class="col-md-3 mb-4">
                <div class="card h-100 shadow-sm border-0">
                    <div class="card-header text-center bg-light text-dark border-bottom-0 py-2">
                        <h4 class="m-0 fw-bold" style="color: #1F4E78;">{year}</h4>
                    </div>
                    <div class="card-body text-center">
                        <div class="mb-3">
                            <div class="text-muted small text-uppercase" style="font-size: 0.75em;">Total Leads</div>
                            <div class="fs-4 fw-bold text-dark">{int(curr['Total']):,.0f}{delta_leads_str}</div>
                        </div>
                        <div class="mb-3">
                            <div class="text-muted small text-uppercase" style="font-size: 0.75em;">Matrículas</div>
                            <div class="fs-4 fw-bold text-success">{int(curr['Matriculas']):,.0f}{delta_mat_str}</div>
                        </div>
                        <div>
                            <div class="text-muted small text-uppercase" style="font-size: 0.75em;">Conversão</div>
                            <div class="fs-5 fw-bold text-primary">{conv_rate:.1f}%</div>
                        </div>
                    </div>
                </div>
            </div>
            """

    # Gerar Cards Mensais com Delta YoY (Scorecards)
    monthly_section_html = ""
    if not timeline_data.empty:
        # Ordenar reverso (mais recente primeiro)
        sorted_months = sorted(timeline_data.index.tolist(), reverse=True)
        
        # Preparar Dropdown de Meses
        months_pt = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
                     'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        
        dropdown_opts = '<option value="" selected disabled>Selecione um mês...</option>'
        dropdown_opts += '<option value="all">Ver Todos</option>'
        for i, m_name in enumerate(months_pt, 1):
            dropdown_opts += f'<option value="{i}">{m_name}</option>'
        
        cards_html = ""
        for month_str in sorted_months:
            curr = timeline_data.loc[month_str]
            m_int = int(month_str.split('-')[1])
            
            # Calcular mês do ano anterior
            try:
                year = int(month_str[:4])
                month = int(month_str[5:])
                prev_year_month_str = f"{year-1}-{month:02d}"
            except:
                prev_year_month_str = ""
            
            delta_leads_str = ""
            delta_mat_str = ""
            
            if prev_year_month_str in timeline_data.index:
                prev = timeline_data.loc[prev_year_month_str]
                
                # Leads Delta
                if prev['Total'] > 0:
                    d_leads = ((curr['Total'] - prev['Total']) / prev['Total']) * 100
                    color = "text-success" if d_leads >= 0 else "text-danger"
                    icon = "▲" if d_leads >= 0 else "▼"
                    delta_leads_str = f'<span class="{color} ms-1" style="font-size: 0.7em;">{icon} {d_leads:.1f}%</span>'
                
                # Matriculas Delta
                if prev['Matriculas'] > 0:
                    d_mat = ((curr['Matriculas'] - prev['Matriculas']) / prev['Matriculas']) * 100
                    color = "text-success" if d_mat >= 0 else "text-danger"
                    icon = "▲" if d_mat >= 0 else "▼"
                    delta_mat_str = f'<span class="{color} ms-1" style="font-size: 0.7em;">{icon} {d_mat:.1f}%</span>'
            
            display_date = format_date_pt(month_str)
            conv_rate = (curr['Matriculas'] / curr['Total'] * 100) if curr['Total'] > 0 else 0
            
            cards_html += f"""
            <div class="col-xl-2 col-lg-3 col-md-4 col-sm-6 mb-3 month-card" data-month="{m_int}" style="display: none;">
                <div class="card h-100 shadow-sm border-0">
                    <div class="card-header text-center bg-white border-bottom py-1">
                        <strong class="text-secondary" style="font-size: 0.9em;">{display_date}</strong>
                    </div>
                    <div class="card-body p-2 text-center">
                        <div class="mb-2">
                            <div class="text-muted small text-uppercase" style="font-size: 0.65em;">Total Leads</div>
                            <div class="fw-bold text-dark">{int(curr['Total']):,.0f}{delta_leads_str}</div>
                        </div>
                        <div class="mb-2">
                            <div class="text-muted small text-uppercase" style="font-size: 0.65em;">Matrículas</div>
                            <div class="fw-bold text-success">{int(curr['Matriculas']):,.0f}{delta_mat_str}</div>
                        </div>
                        <div class="row g-0 border-top pt-2 mt-2">
                            <div class="col-6 border-end">
                                <div class="text-muted small text-uppercase" style="font-size: 0.6em;">Conv.</div>
                                <div class="fw-bold text-primary" style="font-size: 0.85em;">{conv_rate:.1f}%</div>
                            </div>
                            <div class="col-6">
                                <div class="text-muted small text-uppercase" style="font-size: 0.6em;">Perd.</div>
                                <div class="fw-bold text-danger" style="font-size: 0.85em;">{int(curr['Perdidos']):,.0f}</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            """

        monthly_section_html = f"""
        <div class="col-12 mt-4 mb-3">
            <div class="d-flex justify-content-between align-items-center border-bottom pb-2">
                <h5 class="mb-0 text-secondary">Detalhamento Mensal (Comparativo YoY)</h5>
                <div class="d-flex align-items-center">
                    <label for="monthSelect" class="me-2 fw-bold text-secondary">Mês:</label>
                    <select id="monthSelect" class="form-select form-select-sm" style="width: 200px;">
                        {dropdown_opts}
                    </select>
                </div>
            </div>
        </div>
        {cards_html}
        """

    # Gerar Cards por Ciclo Comercial
    cycle_section_html = ""
    if not timeline_data.empty:
        # Agregação por Ciclo
        cycles = {}
        for month_str in timeline_data.index:
            y = int(month_str[:4])
            m = int(month_str[5:])
            
            # Lógica do Ciclo: Out-Mar = AnoSeguinte.1 (Alta), Abr-Set = AnoAtual.2 (Baixa)
            if m >= 10:
                c_code = f"{y+1}.1"
                c_type = "Alta"
            elif m <= 3:
                c_code = f"{y}.1"
                c_type = "Alta"
            else:
                c_code = f"{y}.2"
                c_type = "Baixa"
                
            if c_code not in cycles:
                cycles[c_code] = {'Total': 0, 'Matriculas': 0, 'Perdidos': 0, 'Type': c_type}
            
            curr = timeline_data.loc[month_str]
            cycles[c_code]['Total'] += curr['Total']
            cycles[c_code]['Matriculas'] += curr['Matriculas']
            cycles[c_code]['Perdidos'] += curr['Perdidos']
            
        sorted_cycles = sorted(cycles.keys(), reverse=True)
        
        cycle_cards = ""
        for c_code in sorted_cycles:
            curr = cycles[c_code]
            c_type = curr['Type']
            
            # Comparação Homóloga (ex: 25.1 vs 24.1)
            try:
                cy_year = int(c_code.split('.')[0])
                cy_suffix = c_code.split('.')[1]
                prev_code = f"{cy_year-1}.{cy_suffix}"
            except:
                prev_code = ""
                
            delta_leads_str = ""
            delta_mat_str = ""
            
            if prev_code in cycles:
                prev = cycles[prev_code]
                if prev['Total'] > 0:
                    d_leads = ((curr['Total'] - prev['Total']) / prev['Total']) * 100
                    color = "text-success" if d_leads >= 0 else "text-danger"
                    icon = "▲" if d_leads >= 0 else "▼"
                    delta_leads_str = f'<span class="{color} ms-1" style="font-size: 0.7em;">{icon} {d_leads:.1f}%</span>'
                
                if prev['Matriculas'] > 0:
                    d_mat = ((curr['Matriculas'] - prev['Matriculas']) / prev['Matriculas']) * 100
                    color = "text-success" if d_mat >= 0 else "text-danger"
                    icon = "▲" if d_mat >= 0 else "▼"
                    delta_mat_str = f'<span class="{color} ms-1" style="font-size: 0.7em;">{icon} {d_mat:.1f}%</span>'

            conv_rate = (curr['Matriculas'] / curr['Total'] * 100) if curr['Total'] > 0 else 0
            
            cycle_cards += f"""
            <div class="col-xl-3 col-lg-4 col-md-6 mb-3 cycle-card" data-type="{c_type}" style="display: none;">
                <div class="card h-100 shadow-sm border-0" style="border-top: 3px solid {'#C00000' if c_type == 'Alta' else '#1F4E78'} !important;">
                    <div class="card-header text-center bg-white border-bottom py-2">
                        <strong class="text-dark" style="font-size: 1.1em;">Ciclo {c_code}</strong>
                        <span class="badge {'bg-danger' if c_type == 'Alta' else 'bg-primary'} ms-2" style="font-size: 0.7em;">{c_type}</span>
                    </div>
                    <div class="card-body p-3 text-center">
                        <div class="row mb-2">
                            <div class="col-6 border-end">
                                <div class="text-muted small text-uppercase" style="font-size: 0.65em;">Leads</div>
                                <div class="fw-bold text-dark">{int(curr['Total']):,.0f}</div>
                                <div>{delta_leads_str}</div>
                            </div>
                            <div class="col-6">
                                <div class="text-muted small text-uppercase" style="font-size: 0.65em;">Matrículas</div>
                                <div class="fw-bold text-success">{int(curr['Matriculas']):,.0f}</div>
                                <div>{delta_mat_str}</div>
                            </div>
                        </div>
                        <div class="pt-2 border-top">
                             <div class="text-muted small text-uppercase" style="font-size: 0.65em;">Conversão</div>
                             <div class="fw-bold text-primary">{conv_rate:.1f}%</div>
                        </div>
                    </div>
                </div>
            </div>
            """
            
        cycle_section_html = f"""
        <div class="col-12 mt-4 mb-3">
            <div class="d-flex justify-content-between align-items-center border-bottom pb-2">
                <h5 class="mb-0 text-secondary">Performance por Ciclo Comercial</h5>
                <div class="d-flex align-items-center">
                    <label for="cycleSelect" class="me-2 fw-bold text-secondary">Ciclo:</label>
                    <select id="cycleSelect" class="form-select form-select-sm" style="width: 200px;">
                        <option value="" selected disabled>Selecione...</option>
                        <option value="all">Ver Todos</option>
                        <option value="Alta">Alta (Out-Mar)</option>
                        <option value="Baixa">Baixa (Abr-Set)</option>
                    </select>
                </div>
            </div>
        </div>
        {cycle_cards}
        """

    # Gerar Seção de Canais (Geral + Por Unidade)
    channel_section_html = ""
    if not analise_canais.empty:
        # 1. Visão Geral (Top Cards)
        top_channels_cards = ""
        for channel, row in analise_canais.head(8).iterrows(): # Top 8
            total = int(row['Total'])
            mat = int(row['Matriculas'])
            conv = (mat / total * 100) if total > 0 else 0
            
            top_channels_cards += f"""
            <div class="col-xl-3 col-lg-4 col-md-6 mb-3">
                <div class="card h-100 shadow-sm border-0 bg-light">
                    <div class="card-body p-3">
                        <h6 class="card-title text-truncate" title="{channel}" style="color: #1F4E78;">{channel}</h6>
                        <div class="d-flex justify-content-between align-items-end mt-2">
                            <div>
                                <div class="small text-muted">Leads</div>
                                <div class="fw-bold">{total:,.0f}</div>
                            </div>
                            <div class="text-end">
                                <div class="small text-muted">Conv.</div>
                                <div class="fw-bold text-success">{conv:.1f}%</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            """
            
        # 2. Por Unidade
        unit_channel_html = ""
        uc_type_opts = ""
        uc_opts = ""
        
        if not analise_canais_unidade.empty:
            # Dropdown options
            types_channels = sorted(analise_canais_unidade['Tipo'].unique().tolist())
            units_channels = sorted(analise_canais_unidade['Unidade Desejada'].unique().tolist())
            
            uc_type_opts = '<option value="all">Todos</option>'
            for t in types_channels:
                uc_type_opts += f'<option value="{t}">{t}</option>'
            
            uc_opts = '<option value="" selected disabled>Selecione uma unidade...</option>'
            uc_opts += '<option value="all">Ver Todas</option>'
            for u in units_channels:
                uc_opts += f'<option value="{u}">{u}</option>'
                
            # Cards - Agrupados por Unidade
            acu_sorted = analise_canais_unidade.sort_values(['Unidade Desejada', 'Total'], ascending=[True, False])
            unique_units = acu_sorted['Unidade Desejada'].unique()
            
            for u in unique_units:
                unit_df = acu_sorted[acu_sorted['Unidade Desejada'] == u]
                if unit_df.empty: continue
                
                t = unit_df.iloc[0]['Tipo']
                cards_html = ""
                
                for _, row in unit_df.iterrows():
                    c = row['Canal']
                    total = int(row['Total'])
                    mat = int(row['Matriculas'])
                    conv = (mat / total * 100) if total > 0 else 0
                    
                    cards_html += f"""
                    <div class="col-xl-3 col-lg-4 col-md-6 mb-3">
                        <div class="card h-100 shadow-sm border-0">
                            <div class="card-body p-3">
                                <div class="d-flex justify-content-between mb-2">
                                    <strong class="text-truncate" title="{c}" style="max-width: 70%;">{c}</strong>
                                    <span class="badge bg-light text-dark border">{total}</span>
                                </div>
                                <div class="progress" style="height: 6px;">
                                    <div class="progress-bar bg-success" role="progressbar" style="width: {conv}%" aria-valuenow="{conv}" aria-valuemin="0" aria-valuemax="100"></div>
                                </div>
                                <div class="d-flex justify-content-between mt-1">
                                    <small class="text-muted">Matrículas: {mat}</small>
                                    <small class="text-success fw-bold">{conv:.1f}%</small>
                                </div>
                            </div>
                        </div>
                    </div>
                    """
                
                unit_channel_html += f"""
                <div class="channel-unit-group mb-4" data-type="{t}" data-unit="{u}" style="display: none;">
                    <h6 class="border-bottom pb-2 mb-3" style="color: #1F4E78; font-weight: bold;">
                        {u} <span class="badge {'bg-danger' if t == 'Própria' else 'bg-primary'} ms-2" style="font-size: 0.7em;">{t}</span>
                    </h6>
                    <div class="row">
                        {cards_html}
                    </div>
                </div>
                """
        
        channel_section_html = f"""
        <div class="col-12 mt-4 mb-3">
            <h5 class="mb-3 text-secondary border-bottom pb-2">Análise de Canais (Geral)</h5>
            <div class="row">{top_channels_cards}</div>
            
            <div class="d-flex justify-content-between align-items-center border-bottom pb-2 mt-4 mb-3 flex-wrap">
                <h5 class="mb-0 text-secondary">Canais por Unidade</h5>
                <div class="d-flex align-items-center mt-2 mt-sm-0">
                    <div class="me-3">
                        <label for="channelUnitTypeSelect" class="me-1 fw-bold text-secondary small">Tipo:</label>
                        <select id="channelUnitTypeSelect" class="form-select form-select-sm d-inline-block" style="width: 120px;">
                            {uc_type_opts}
                        </select>
                    </div>
                    <div>
                        <label for="channelUnitSelect" class="me-1 fw-bold text-secondary small">Unidade:</label>
                        <select id="channelUnitSelect" class="form-select form-select-sm d-inline-block" style="width: 200px;">
                            {uc_opts}
                        </select>
                    </div>
                </div>
            </div>
            <div id="unitChannelsContainer">{unit_channel_html}</div>
        </div>
        """

    # Gerar Cards de Unidades (Análise de Unidades)
    unit_section_html = ""
    if not analise.empty:
        # Ordenar por Total de Leads
        analise_sorted = analise.sort_values('Total', ascending=False)
        
        # Preparar Dropdowns
        types = sorted(analise.index.get_level_values('Tipo').unique().tolist())
        units = sorted(analise.index.get_level_values('Unidade Desejada').unique().tolist())
        
        type_opts = '<option value="all">Todos</option>'
        for t in types:
            type_opts += f'<option value="{t}">{t}</option>'
            
        unit_opts = '<option value="" selected disabled>Selecione uma unidade...</option>'
        unit_opts += '<option value="all">Ver Todas</option>'
        for u in units:
            unit_opts += f'<option value="{u}">{u}</option>'
            
        unit_cards = ""
        for (tipo, unidade), row in analise_sorted.iterrows():
            total = int(row['Total'])
            mat = int(row['Matriculas'])
            perd = int(row['Perdidos'])
            conv = row['Taxa_Conv']
            ativ = row['Ativ_Media_Total']
            
            # Cor da borda baseada no tipo
            border_color = '#C00000' if tipo == 'Própria' else '#1F4E78'
            badge_class = 'bg-danger' if tipo == 'Própria' else 'bg-primary'
            
            unit_cards += f"""
            <div class="col-xl-3 col-lg-4 col-md-6 mb-3 unit-card" data-type="{tipo}" data-unit="{unidade}" style="display: none;">
                <div class="card h-100 shadow-sm border-0" style="border-left: 5px solid {border_color} !important;">
                    <div class="card-header bg-white border-bottom py-2 d-flex justify-content-between align-items-center">
                        <strong class="text-dark text-truncate" style="font-size: 0.95em;" title="{unidade}">{unidade}</strong>
                        <span class="badge {badge_class} ms-1" style="font-size: 0.6em;">{tipo}</span>
                    </div>
                    <div class="card-body p-3">
                        <div class="row text-center mb-2">
                            <div class="col-4 border-end">
                                <div class="text-muted small text-uppercase" style="font-size: 0.6em;">Leads</div>
                                <div class="fw-bold text-dark">{total:,.0f}</div>
                            </div>
                            <div class="col-4 border-end">
                                <div class="text-muted small text-uppercase" style="font-size: 0.6em;">Matr.</div>
                                <div class="fw-bold text-success">{mat:,.0f}</div>
                            </div>
                            <div class="col-4">
                                <div class="text-muted small text-uppercase" style="font-size: 0.6em;">Conv.</div>
                                <div class="fw-bold text-primary">{conv:.1f}%</div>
                            </div>
                        </div>
                        <div class="row text-center border-top pt-2">
                             <div class="col-6 border-end">
                                <div class="text-muted small text-uppercase" style="font-size: 0.6em;">Perdidos</div>
                                <div class="fw-bold text-danger">{perd:,.0f}</div>
                             </div>
                             <div class="col-6">
                                <div class="text-muted small text-uppercase" style="font-size: 0.6em;">Ativ. Méd.</div>
                                <div class="fw-bold text-secondary">{ativ:.1f}</div>
                             </div>
                        </div>
                    </div>
                </div>
            </div>
            """
            
        unit_section_html = f"""
        <div class="col-12 mt-4 mb-3">
            <div class="d-flex justify-content-between align-items-center border-bottom pb-2 flex-wrap">
                <h5 class="mb-0 text-secondary">Análise de Unidades</h5>
                <div class="d-flex align-items-center mt-2 mt-sm-0">
                    <div class="me-3">
                        <label for="unitTypeSelect" class="me-1 fw-bold text-secondary small">Tipo:</label>
                        <select id="unitTypeSelect" class="form-select form-select-sm d-inline-block" style="width: 120px;">
                            {type_opts}
                        </select>
                    </div>
                    <div>
                        <label for="unitSelect" class="me-1 fw-bold text-secondary small">Unidade:</label>
                        <select id="unitSelect" class="form-select form-select-sm d-inline-block" style="width: 200px;">
                            {unit_opts}
                        </select>
                    </div>
                </div>
            </div>
        </div>
        {unit_cards}
        """
            
    # HTML Template
    html_content = f"""
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Dashboard Performance - Red Balloon</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <style>
            body {{ background-color: #f8f9fa; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }}
            .navbar {{ background-color: #C00000 !important; }}
            .card {{ margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border: none; border-radius: 8px; }}
            .card-header {{ background-color: #1F4E78; color: white; font-weight: bold; border-radius: 8px 8px 0 0 !important; }}
            .kpi-card {{ text-align: center; padding: 15px; border-radius: 8px; background: white; height: 100%; border-left: 5px solid #C00000; }}
            .kpi-value {{ font-size: 1.6em; font-weight: bold; color: #1F4E78; }}
            .kpi-sub {{ font-size: 0.8em; color: #999; }}
            .kpi-label {{ color: #6c757d; font-size: 0.85em; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 5px; }}
            .table-responsive {{ max-height: 600px; overflow-y: auto; }}
            thead th {{ position: sticky; top: 0; background-color: #1F4E78 !important; color: white; z-index: 1; }}
            .table-striped tbody tr:nth-of-type(odd) {{ background-color: rgba(0,0,0,.02); }}
        </style>
    </head>
    <body>
        <nav class="navbar navbar-dark mb-4">
            <div class="container-fluid">
                <span class="navbar-brand mb-0 h1">RED BALLOON | Dashboard de Performance</span>
                <span class="text-light" style="font-size: 0.9em">Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}</span>
            </div>
        </nav>

        <div class="container-fluid px-4">
            <!-- KPIs -->
            <div class="row mb-4">
    """
    
    for label, val_valid, val_raw, fmt in summary_data:
        fmt_valid = f"{val_valid:,.0f}".replace(",", ".") if fmt == '#,##0' else f"{val_valid*100:.1f}%"
        fmt_raw = f"{val_raw:,.0f}".replace(",", ".") if fmt == '#,##0' else f"{val_raw*100:.1f}%"
        
        html_content += f"""
                <div class="col-lg-2 col-md-4 col-sm-6 mb-3">
                    <div class="kpi-card shadow-sm">
                        <div class="kpi-label">{label}</div>
                        <div class="kpi-value">{fmt_valid}</div>
                        <div class="kpi-sub" title="Cenário Total (Bruto)">Total: {fmt_raw}</div>
                    </div>
                </div>
        """
    
    html_content += f"""
            </div>

            <!-- Gráficos -->
            <div class="row mb-4">
                <div class="col-lg-8 mb-4">
                    <div class="card h-100">
                        <div class="card-header">Evolução Temporal (Leads vs Matrículas vs Perdidos)</div>
                        <div class="card-body">
                            <div style="width: 100%; overflow-x: auto;">
                                <div style="position: relative; height: 300px; min-width: {timeline_min_width}px;">
                                    <canvas id="timelineChart"></canvas>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="col-lg-4 mb-4">
                    <div class="card h-100">
                        <div class="card-header">Top 10 Canais (Volume)</div>
                        <div class="card-body">
                            <div style="position: relative; height: 300px; width: 100%;">
                                <canvas id="channelsChart"></canvas>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Comparativo Anual -->
            <div class="row">
                <div class="col-12">
                    <h5 class="mb-3 text-secondary border-bottom pb-2">Comparativo Anual (Score Points)</h5>
                </div>
                {yearly_cards_html}
                {monthly_section_html}
                {cycle_section_html}
                {channel_section_html}
                {unit_section_html}
            </div>
            
            <div class="row">
                <div class="col-12 text-center text-muted small mt-4 mb-4">
                    <p>Relatório Macro - Red Balloon</p>
                </div>
            </div>
        </div>

        <script>
            const ctxTimeline = document.getElementById('timelineChart').getContext('2d');
            new Chart(ctxTimeline, {{
                type: 'line',
                data: {{
                    labels: {json.dumps(timeline_labels)},
                    datasets: [
                        {{ label: 'Total Leads', data: {json.dumps(timeline_leads)}, borderColor: '#1F4E78', backgroundColor: 'rgba(31, 78, 120, 0.1)', tension: 0.3, fill: true }},
                        {{ label: 'Matrículas', data: {json.dumps(timeline_matriculas)}, borderColor: '#70AD47', backgroundColor: 'rgba(112, 173, 71, 0.1)', tension: 0.3, fill: true }},
                        {{ label: 'Perdidos', data: {json.dumps(timeline_perdidos)}, borderColor: '#C00000', backgroundColor: 'rgba(192, 0, 0, 0.05)', tension: 0.3, fill: true }}
                    ]
                }},
                options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'top' }} }} }}
            }});

            const ctxChannels = document.getElementById('channelsChart').getContext('2d');
            new Chart(ctxChannels, {{
                type: 'bar',
                data: {{
                    labels: {json.dumps(canais_labels)},
                    datasets: [
                        {{ label: 'Total Leads', data: {json.dumps(canais_leads)}, backgroundColor: '#1F4E78' }},
                        {{ label: 'Matrículas', data: {json.dumps(canais_matriculas)}, backgroundColor: '#70AD47' }}
                    ]
                }},
                options: {{ indexAxis: 'y', responsive: true, maintainAspectRatio: false }}
            }});

            // Filtro de Mês
            const monthSelect = document.getElementById('monthSelect');
            if(monthSelect) {{
                monthSelect.addEventListener('change', function() {{
                    const val = this.value;
                    const cards = document.querySelectorAll('.month-card');
                    cards.forEach(c => {{
                        if (val === 'all') {{
                            c.style.display = '';
                        }} else if (c.getAttribute('data-month') === val) {{
                            c.style.display = '';
                        }} else {{
                            c.style.display = 'none';
                        }}
                    }});
                }});
            }}
            
            // Filtro de Ciclo
            const cycleSelect = document.getElementById('cycleSelect');
            if(cycleSelect) {{
                cycleSelect.addEventListener('change', function() {{
                    const val = this.value;
                    const cards = document.querySelectorAll('.cycle-card');
                    cards.forEach(c => {{
                        if (val === 'all') {{
                            c.style.display = '';
                        }} else if (c.getAttribute('data-type') === val) {{
                            c.style.display = '';
                        }} else {{
                            c.style.display = 'none';
                        }}
                    }});
                }});
            }}
            
            // Filtro de Unidades
            const unitTypeSelect = document.getElementById('unitTypeSelect');
            const unitSelect = document.getElementById('unitSelect');
            
            if(unitTypeSelect && unitSelect) {{
                function filterUnits() {{
                    const typeVal = unitTypeSelect.value;
                    const unitVal = unitSelect.value;
                    const cards = document.querySelectorAll('.unit-card');
                    
                    cards.forEach(c => {{
                        const cType = c.getAttribute('data-type');
                        const cUnit = c.getAttribute('data-unit');
                        
                        let show = true;
                        if (typeVal !== 'all' && cType !== typeVal) show = false;
                        if (unitVal === '') show = false; // Ocultar se nada selecionado
                        else if (unitVal !== 'all' && cUnit !== unitVal) show = false;
                        
                        c.style.display = show ? '' : 'none';
                    }});
                }}
                
                unitTypeSelect.addEventListener('change', filterUnits);
                unitSelect.addEventListener('change', filterUnits);
            }}
            
            // Filtro de Canais por Unidade
            const channelUnitTypeSelect = document.getElementById('channelUnitTypeSelect');
            const channelUnitSelect = document.getElementById('channelUnitSelect');
            
            if(channelUnitTypeSelect && channelUnitSelect) {{
                function filterChannelUnits() {{
                    const typeVal = channelUnitTypeSelect.value;
                    const unitVal = channelUnitSelect.value;
                    const groups = document.querySelectorAll('.channel-unit-group');
                    
                    groups.forEach(g => {{
                        const cType = g.getAttribute('data-type');
                        const cUnit = g.getAttribute('data-unit');
                        
                        let show = true;
                        if (typeVal !== 'all' && cType !== typeVal) show = false;
                        if (unitVal === '') show = false;
                        else if (unitVal !== 'all' && cUnit !== unitVal) show = false;
                        
                        if (show) {{
                            g.style.display = '';
                        }} else {{
                            g.style.display = 'none';
                        }}
                    }});
                }}
                
                channelUnitTypeSelect.addEventListener('change', filterChannelUnits);
                channelUnitSelect.addEventListener('change', filterChannelUnits);
            }}
        </script>
    </body>
    </html>
    """
    
    with open(html_file, "w", encoding="utf-8") as f:
        f.write(html_content)
    
    return html_file

if __name__ == "__main__":
    print("=" * 100)
    print("GERADOR DE DASHBOARD HTML - RED BALLOON")
    print("=" * 100)
    
    try:
        summary, analise, timeline, canais, yearly, canais_unidade = process_data()
        file_path = generate_html(summary, analise, timeline, canais, yearly, canais_unidade)
        
        print("\n✅ DASHBOARD GERADO COM SUCESSO!")
        print(f"📂 Arquivo: {file_path}")
        print("=" * 100)
        
    except Exception as e:
        print(f"\n❌ ERRO: {str(e)}")
        logger.error(str(e), exc_info=True)
        sys.exit(1)