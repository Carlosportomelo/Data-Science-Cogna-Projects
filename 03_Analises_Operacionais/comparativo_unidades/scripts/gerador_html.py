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
    if 'dt_create' in df.columns:
        df_timeline = df[df['dt_create'].notna()].copy()
        df_timeline['Periodo'] = df_timeline['dt_create'].dt.to_period('M')
        timeline_data = df_timeline.groupby('Periodo').agg({
            'Convertido': 'sum',
            'Perdido': 'sum'
        }).rename(columns={'Convertido': 'Matriculas', 'Perdido': 'Perdidos'})
        timeline_data.index = timeline_data.index.astype(str)

    # Canais
    analise_canais = pd.DataFrame()
    col_fonte = None
    keywords = ['fonte original', 'original source', 'source', 'fonte', 'origem']
    for col in df.columns:
        if any(k in col.lower() for k in keywords):
            col_fonte = col
            break
    
    if col_fonte:
        df[col_fonte] = df[col_fonte].fillna('Não Identificado').astype(str).str.strip()
        analise_canais = df.groupby(col_fonte).agg({
            'Record ID': 'count',
            'Convertido': 'sum'
        }).rename(columns={'Record ID': 'Total', 'Convertido': 'Matriculas'}).sort_values('Total', ascending=False)

    summary_data = [
        ("Total de Leads", total, raw_total, '#,##0'),
        ("Matrículas", conv, raw_conv, '#,##0'),
        ("Taxa de Conversão", conv / total if total > 0 else 0, raw_conv / raw_total if raw_total > 0 else 0, '0.0%'),
        ("Perdidos", perd, raw_perd, '#,##0'),
        ("Taxa de Perda", perd / total if total > 0 else 0, raw_perd / raw_total if raw_total > 0 else 0, '0.0%'),
        ("Em Pipeline", total - conv - perd, raw_total - raw_conv - raw_perd, '#,##0')
    ]

    return summary_data, analise, timeline_data, analise_canais

def generate_html(summary_data, analise, timeline_data, analise_canais):
    """Gera o arquivo HTML."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    html_file = OUTPUT_DIR / f"DASHBOARD_PERFORMANCE_{timestamp}.html"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Preparar dados JSON
    timeline_labels = timeline_data.index.tolist() if not timeline_data.empty else []
    timeline_matriculas = timeline_data['Matriculas'].tolist() if not timeline_data.empty else []
    timeline_perdidos = timeline_data['Perdidos'].tolist() if not timeline_data.empty else []
    
    if not analise_canais.empty:
        top_canais = analise_canais.head(10)
        canais_labels = top_canais.index.tolist()
        canais_matriculas = top_canais['Matriculas'].tolist()
        canais_leads = top_canais['Total'].tolist()
    else:
        canais_labels, canais_matriculas, canais_leads = [], [], []

    # Tabela HTML
    analise_html = analise.reset_index().copy()
    analise_html = analise_html.replace([float('inf'), -float('inf')], 0).fillna(0)
    cols_rename = {
        'Unidade Desejada': 'Unidade', 'Total': 'Leads', 'Matriculas': 'Matr.', 'Perdidos': 'Perd.',
        'Taxa_Conv': '% Conv.', 'Taxa_Perda': '% Perda', 'Leads_p_Matricula': 'Leads/Matr.',
        'Ativ_Media_Total': 'Ativ. Geral', 'Ativ_Media_Mat': 'Ativ. Matr.', 'Ativ_Media_Perd': 'Ativ. Perd.'
    }
    analise_html = analise_html.rename(columns=cols_rename)
    
    # Formatação
    for col in analise_html.columns:
        if 'Ativ.' in col or 'Leads/Matr.' in col:
            analise_html[col] = analise_html[col].apply(lambda x: f"{x:.1f}")
        elif '%' in col:
            analise_html[col] = analise_html[col].apply(lambda x: f"{x:.1f}%")
            
    table_html = analise_html.to_html(classes="table table-hover table-sm", index=False, border=0, float_format=lambda x: f"{x:,.0f}".replace(",", "."))

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
                <div class="col-md-2 col-sm-4 mb-3">
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
                        <div class="card-header">Evolução Temporal (Matrículas vs Perdidos)</div>
                        <div class="card-body">
                            <div style="position: relative; height: 300px; width: 100%;">
                                <canvas id="timelineChart"></canvas>
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

            <!-- Tabela Detalhada -->
            <div class="row">
                <div class="col-12">
                    <div class="card">
                        <div class="card-header d-flex justify-content-between align-items-center">
                            <span>Detalhamento Completo por Unidade</span>
                            <span class="badge bg-light text-dark">Total: {len(analise_html)} Unidades</span>
                        </div>
                        <div class="card-body table-responsive p-0">
                            {table_html}
                        </div>
                    </div>
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
        summary, analise, timeline, canais = process_data()
        file_path = generate_html(summary, analise, timeline, canais)
        
        print("\n✅ DASHBOARD GERADO COM SUCESSO!")
        print(f"📂 Arquivo: {file_path}")
        print("=" * 100)
        
    except Exception as e:
        print(f"\n❌ ERRO: {str(e)}")
        logger.error(str(e), exc_info=True)
```

### Como usar:
Basta executar este novo script (`gerador_dashboard.py`) no seu terminal. Ele encontrará automaticamente o arquivo CSV mais recente (usando a mesma lógica do script original), processará os dados e salvará o arquivo HTML na pasta `outputs`.

<!--
[PROMPT_SUGGESTION]Como posso adicionar um filtro de busca na tabela HTML para encontrar rapidamente uma unidade específica?[/PROMPT_SUGGESTION]
[PROMPT_SUGGESTION]Poderia adicionar um botão no HTML para exportar a tabela filtrada de volta para CSV?[/PROMPT_SUGGESTION]
->