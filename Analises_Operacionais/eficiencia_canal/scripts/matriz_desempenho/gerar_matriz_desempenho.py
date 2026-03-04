import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
from datetime import datetime
import matplotlib.ticker as mtick
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# --- CONFIGURAÇÃO DE UNIDADES ---
# Adicione aqui os nomes exatos das unidades próprias para a classificação
UNIDADES_PROPRIAS = [
    "VILA LEOPOLDINA",
    "MORUMBI",
    "PACAEMBU",
    "ITAIM BIBI",
    "PINHEIROS",
    "JARDINS",
    "PERDIZES",
    "SANTANA",
    "MOOCA"
]

def gerar_matriz_unidades():
    # --- CONFIGURAÇÕES ---
    project_root = Path(r'c:\Users\a483650\Projetos\analise_eficiencia_canal')
    
    # Busca o relatório consolidado mais recente
    outputs_dir = project_root / 'outputs'
    report_dir = outputs_dir / 'analise_consolidada_leads_matriculas'
    
    try:
        files = list(report_dir.glob('analise_consolidada_leads_matriculas_*.xlsx'))
        if not files:
            print(f"ERRO: Nenhum arquivo consolidado encontrado em {report_dir}")
            return
        data_file = max(files, key=lambda f: f.stat().st_mtime)
    except Exception as e:
        print(f"Erro ao buscar arquivo: {e}")
        return

    output_dir = outputs_dir / 'matriz_desempenho'
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Data do Relatório
    today_str = datetime.now().strftime('%Y-%m-%d')
    output_img = output_dir / f'matriz_desempenho_unidades_{today_str}.png'
    output_html = output_dir / f'matriz_desempenho_interativa_{today_str}.html'

    print(f"--- GERANDO MATRIZ DE DESEMPENHO ---")
    print(f"Lendo arquivo consolidado: {data_file.name}")

    # --- 1. LEITURA E LIMPEZA ---
    try:
        # Lê as abas específicas do relatório consolidado
        df_leads_raw = pd.read_excel(data_file, sheet_name='Leads (Unidades)')
        df_matr_raw = pd.read_excel(data_file, sheet_name='Matrículas (Unidades)')
    except Exception as e:
        print(f"Erro ao ler Excel: {e}")
        print("Verifique se o arquivo possui as abas 'Leads (Unidades)' e 'Matrículas (Unidades)'")
        return

    # --- 2. PROCESSAMENTO ---
    # Define o ciclo alvo (26.1)
    target_cycle = '26.1'
    
    if target_cycle not in df_leads_raw.columns:
        print(f"AVISO: Coluna '{target_cycle}' não encontrada. Tentando identificar...")
        cols = [c for c in df_leads_raw.columns if isinstance(c, str) and '.1' in c]
        if cols:
            target_cycle = cols[-1] # Pega o último ciclo disponível
            print(f"   > Usando ciclo: {target_cycle}")
        else:
            print("ERRO: Ciclo não identificado nas colunas.")
            return

    # --- DEBUG: Detalhamento para SANTANA ---
    print(f"\n[DEBUG] Verificando dados brutos para 'SANTANA' no ciclo {target_cycle}:")
    # Filtra linhas que contenham SANTANA (case insensitive)
    santana_rows = df_matr_raw[df_matr_raw['Unidade'].astype(str).str.upper().str.contains("SANTANA", na=False)]
    if not santana_rows.empty:
        print(santana_rows[['Unidade', 'Fonte original do tráfego', target_cycle]].to_string())
        print(f"Total calculado para Santana: {santana_rows[target_cycle].sum()}\n")

    # Agrupa por Unidade (somando os canais)
    # O relatório consolidado quebra por Unidade + Canal, então precisamos somar
    leads_by_unit = df_leads_raw.groupby('Unidade')[target_cycle].sum().reset_index(name='Leads')
    matr_by_unit = df_matr_raw.groupby('Unidade')[target_cycle].sum().reset_index(name='Matriculas')
    
    # Remove linha de 'TOTAL GERAL' se existir (gerada pelo script anterior)
    leads_by_unit = leads_by_unit[leads_by_unit['Unidade'] != 'TOTAL GERAL']
    matr_by_unit = matr_by_unit[matr_by_unit['Unidade'] != 'TOTAL GERAL']

    # Merge dos dados
    df_final = pd.merge(leads_by_unit, matr_by_unit, on='Unidade', how='outer').fillna(0)
    
    # Normaliza nomes das unidades para o gráfico (Remove espaços e padroniza)
    df_final['Unidade'] = df_final['Unidade'].astype(str).str.strip().str.title()
    # Reagrupa caso a normalização tenha juntado duplicatas (ex: "Santana" e "SANTANA")
    df_final = df_final.groupby('Unidade', as_index=False).sum()

    # --- 2.1 LEITURA DA META (Grafico_dispersao_pronto) ---
    meta_file_xlsx = project_root / 'data' / 'Grafico_dispersao_pronto.xlsx'
    meta_file_csv = project_root / 'data' / 'Grafico_dispersao_pronto.csv'
    
    df_meta = pd.DataFrame()
    if meta_file_xlsx.exists():
        try:
            df_meta = pd.read_excel(meta_file_xlsx)
            print(f"Lendo arquivo de Meta: {meta_file_xlsx.name}")
        except Exception as e:
            print(f"Erro ao ler Meta (xlsx): {e}")
    elif meta_file_csv.exists():
        try:
            df_meta = pd.read_csv(meta_file_csv, sep=';' if ';' in open(meta_file_csv).readline() else ',')
            print(f"Lendo arquivo de Meta: {meta_file_csv.name}")
        except Exception as e:
            print(f"Erro ao ler Meta (csv): {e}")
            
    if not df_meta.empty:
        # Tenta identificar colunas
        cols_lower = {c.lower(): c for c in df_meta.columns}
        col_unidade = next((cols_lower[c] for c in cols_lower if 'unidade' in c or 'campus' in c), None)
        
        # Busca colunas de NOVOS REAL (X) e LEADS (Y)
        # O usuário pediu especificamente "Novo Real" (Coluna C) e "Leads" (Coluna B)
        col_novos_real = next((cols_lower[c] for c in cols_lower if 'novos' in c and 'real' in c), None)
        # Fallback: Tenta achar algo que tenha 'novos' mas não 'meta'
        if not col_novos_real:
            col_novos_real = next((cols_lower[c] for c in cols_lower if 'novos' in c and 'meta' not in c), None)
        
        col_leads_ext = next((cols_lower[c] for c in cols_lower if 'lead' in c), None)
        
        # Busca coluna de META (para cálculo do %)
        col_novos_meta = next((cols_lower[c] for c in cols_lower if 'meta' in c and 'novos' in c), None)
        if not col_novos_meta: col_novos_meta = next((cols_lower[c] for c in cols_lower if 'meta' in c), None)

        if col_unidade and col_novos_real and col_leads_ext and col_novos_meta:
            df_meta = df_meta[[col_unidade, col_novos_real, col_leads_ext, col_novos_meta]].rename(columns={
                col_unidade: 'Unidade', 
                col_novos_real: 'Novos_Real', 
                col_leads_ext: 'Leads_Ext',
                col_novos_meta: 'Novos_Meta'
            })
            df_meta['Unidade'] = df_meta['Unidade'].astype(str).str.strip().str.title()
            df_meta = df_meta.groupby('Unidade', as_index=False).sum()
            
            # Merge com df_final (Outer para incluir unidades que só tenham meta)
            df_final = pd.merge(df_final, df_meta, on='Unidade', how='outer').fillna(0)
            print("Dados de Meta integrados com sucesso.")
        else:
            print("AVISO: Colunas de Unidade/Leads/Novos Real/Meta não identificadas no arquivo. Verifique os nomes.")

    # Remove unidades com volume muito baixo (ruído) se necessário, ou mantém tudo
    # Ajuste: Mantém se tiver Realizado (HubSpot) OU Dados Externos
    if 'Novos_Real' in df_final.columns:
        df_final = df_final[(df_final['Leads'] > 0) | (df_final['Matriculas'] > 0) | (df_final['Novos_Real'] > 0) | (df_final['Leads_Ext'] > 0)]
        df_final['Novos_Real'] = df_final['Novos_Real'].fillna(0)
        df_final['Leads_Ext'] = df_final['Leads_Ext'].fillna(0)
        df_final['Novos_Meta'] = df_final['Novos_Meta'].fillna(0)
        
        # Cálculo do % de Atingimento
        df_final['Atingimento'] = df_final.apply(lambda x: x['Novos_Real'] / x['Novos_Meta'] if x['Novos_Meta'] > 0 else 0, axis=1)
    else:
        df_final = df_final[(df_final['Leads'] > 0) | (df_final['Matriculas'] > 0)]
        df_final['Novos_Real'] = 0
        df_final['Leads_Ext'] = 0

    # Classificação: Própria vs Franqueada
    unidades_proprias_upper = [u.upper() for u in UNIDADES_PROPRIAS]
    df_final['Tipo'] = df_final['Unidade'].apply(lambda x: 'Própria' if str(x).upper() in unidades_proprias_upper else 'Franqueada')
    
    # Calcula Taxa de Conversão (para usar como cor ou tamanho se quiser futuramente)
    df_final['Conversao'] = df_final['Matriculas'] / df_final['Leads']

    print(f"Unidades processadas: {len(df_final)}")
    print(df_final.head())

    # Define colunas padrão para o gráfico (Prioridade: % Atingimento vs Leads)
    use_external = 'Atingimento' in df_final.columns and 'Leads_Ext' in df_final.columns
    x_col = 'Atingimento' if use_external else 'Matriculas'
    y_col = 'Leads_Ext' if use_external else 'Leads'

    # --- 3. GERAÇÃO DO GRÁFICO ---
    plt.figure(figsize=(14, 10))
    sns.set_style("whitegrid")

    # Scatter Plot
    # Padrão: X = Novos Real, Y = Leads (se disponível)
    sns.scatterplot(
        data=df_final, 
        x=x_col, 
        y=y_col, 
        s=150, # Tamanho da bolinha
        alpha=0.7,
        edgecolor='black',
        color='#4C72B0'
    )

    # Adiciona Rótulos (Nomes das Unidades)
    for i in range(df_final.shape[0]):
        plt.text(
            df_final[x_col].iloc[i], 
            df_final[y_col].iloc[i], 
            f"  {df_final.Unidade.iloc[i]}", 
            horizontalalignment='left', 
            verticalalignment='center',
            size='small', 
            color='#333333', 
            weight='semibold'
        )

    # Linhas Médias (Quadrantes)
    avg_y = df_final[y_col].mean()
    avg_x = df_final[x_col].mean()

    plt.axhline(y=avg_y, color='r', linestyle='--', alpha=0.5, label=f'Média Y ({int(avg_y)})')
    plt.axvline(x=avg_x, color='b', linestyle='--', alpha=0.5, label=f'Média X ({avg_x:.1%})' if use_external else f'Média X ({int(avg_x)})')

    # Formatação de Porcentagem no Eixo X
    if use_external:
        plt.gca().xaxis.set_major_formatter(mtick.PercentFormatter(1.0))

    # Títulos e Labels
    title_text = f'Matriz de Desempenho: LEADS vs % ATINGIMENTO META' if use_external else f'Matriz de Desempenho (Ciclo {target_cycle})'
    plt.title(f'{title_text}\nFonte: {data_file.name}', fontsize=16)
    plt.xlabel('% Atingimento da Meta (Real/Meta)' if use_external else 'Matrículas', fontsize=12)
    plt.ylabel('Leads (Coluna B)' if use_external else 'Leads', fontsize=12)
    plt.legend()

    # Salvar
    plt.tight_layout()
    plt.savefig(output_img, dpi=300)
    print(f"Gráfico salvo em: {output_img}")
    
    # --- 4. GERAÇÃO DO HTML INTERATIVO (PLOTLY) ---
    print("Gerando HTML interativo...")
    
    # Garante ordem para os botões funcionarem (Própria primeiro, Franqueada depois)
    df_final = df_final.sort_values('Tipo', ascending=False) 
    
    fig = px.scatter(
        df_final,
        x=x_col,
        y=y_col,
        color='Tipo',
        hover_name='Unidade',
        hover_data={'Matriculas': False, 'Leads': False, 'Conversao': False, 'Tipo': False},
        title=title_text,
        labels={x_col: '% Atingimento' if use_external else 'Matrículas', y_col: 'Leads (Coluna B)' if use_external else 'Leads'},
        color_discrete_map={'Própria': '#1f77b4', 'Franqueada': '#ff7f0e'},
        category_orders={'Tipo': ['Própria', 'Franqueada']} # Força ordem dos traces
    )

    # Define tamanho fixo para as bolinhas (limpeza visual)
    fig.update_traces(marker=dict(size=15, line=dict(width=1, color='DarkSlateGrey')))

    if use_external:
        fig.update_layout(xaxis_tickformat='.0%')

    # Linhas Médias
    fig.add_hline(y=avg_y, line_dash="dash", line_color="gray", annotation_text="Média Leads")
    fig.add_vline(x=avg_x, line_dash="dash", line_color="gray", annotation_text="Média Atingimento")

    # Adiciona dados de Meta ao customdata para troca via botão
    # Estrutura do customdata no Plotly Express: [Matriculas, Leads, Conversao, Tipo, ...] (depende do hover_data)
    # Vamos adicionar arrays auxiliares para os botões usarem

    # --- PREPARAÇÃO PARA O MENU DE UNIDADES ---
    # Captura os dados originais dos traces (0=Própria, 1=Franqueada) para poder restaurar/filtrar
    # Nota: O Plotly Express gera traces baseados na cor. Como ordenamos por Tipo, sabemos:
    # Trace 0 -> Próprias
    # Trace 1 -> Franqueadas
    
    # Garante que temos dados para capturar (pode ser que só tenha 1 tipo)
    traces_data = []
    
    for trace in fig.data:
        t_type = trace.name # 'Própria' ou 'Franqueada'
        traces_data.append({
            'x': trace.x, # Atual (Padrão)
            'y': trace.y, # Atual (Padrão)
            'hovertext': trace.hovertext,
            'type': t_type
        })

    # Botão "Todas" (Restaura tudo)
    unit_buttons = [dict(
        label="Todas as Unidades",
        method="restyle",
        args=[{
            'marker.opacity': [1.0] * len(traces_data),
            'textposition': 'top center' # Restaura labels se necessário
        }]
    )]

    # Botões Individuais por Unidade
    all_units_sorted = sorted(df_final['Unidade'].unique())
    
    for unit in all_units_sorted:
        # Em vez de filtrar X/Y, vamos alterar a OPACIDADE
        # 1 = Visível, 0.1 = Quase invisível (ou 0 para sumir)
        opacity_args = []
        
        for i, t_data in enumerate(traces_data):
            # Cria array de opacidade para este trace
            current_opacity = []
            if t_data['hovertext'] is not None and unit in t_data['hovertext']:
                # Se a unidade está neste trace, marca 1 para ela e 0 para os outros
                for h_text in t_data['hovertext']:
                    current_opacity.append(1.0 if unit in h_text else 0.05)
            else:
                # Trace não contém a unidade, tudo transparente
                current_opacity = [0.05] * len(t_data['x'])
            
            opacity_args.append(current_opacity)
        
        unit_buttons.append(dict(
            label=unit,
            method="restyle",
            args=[{
                'marker.opacity': opacity_args
            }]
        ))

    # Botões de Filtro (Dropdown)
    fig.update_layout(
        updatemenus=[
            # Menu 1: Tipo (Própria/Franqueada)
            dict(
                buttons=list([
                    dict(label="Todas",
                         method="update",
                         args=[{"visible": [True, True]}, # Mostra ambos os traces
                               ]),
                    dict(label="Apenas Próprias",
                         method="update",
                         args=[{"visible": [True, False]}, # Mostra trace 0 (Própria), esconde 1
                               ]),
                    dict(label="Apenas Franqueadas",
                         method="update",
                         args=[{"visible": [False, True]}, # Esconde trace 0, mostra 1 (Franqueada)
                               ]),
                ]),
                direction="down", showactive=True, x=0.0, xanchor="left", y=1.15, yanchor="top"
            ),
            # Menu 2: Unidade Específica
            dict(
                buttons=unit_buttons,
                direction="down", showactive=True, x=0.15, xanchor="left", y=1.15, yanchor="top",
                active=0
            ),
        ]
    )
    
    fig.write_html(output_html)
    print(f"HTML salvo em: {output_html}")

if __name__ == "__main__":
    gerar_matriz_unidades()
