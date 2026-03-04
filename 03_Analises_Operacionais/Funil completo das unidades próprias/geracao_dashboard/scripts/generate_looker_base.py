import pandas as pd
import os
import sys

# Ajusta o encoding do stdout para evitar erros ao imprimir
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass


def generate_looker_base(caminho_hubspot: str, output_path: str, unidades_proprias=None):
    unidades_proprias = unidades_proprias or []

    print(f"Lendo a base de dados do HubSpot: {caminho_hubspot}")
    df = pd.read_csv(caminho_hubspot, encoding='utf-8', low_memory=False)

    # Mapeamento de colunas (usa apenas as existentes)
    colunas_interesse = {
        'Record ID': 'ID_Negocio',
        'Nome do negócio': 'Nome_Negocio',
        'Data de criação': 'Data_Criacao',
        'Etapa do negócio': 'Etapa_Negocio',
        'Fonte original do tráfego': 'Fonte_Trafego',
        'Detalhamento da fonte original do tráfego 1': 'Fonte_Detalhe_1',
        'Detalhamento da fonte original do tráfego 2': 'Fonte_Detalhe_2',
        'Data de fechamento': 'Data_Fechamento',
        'Unidade Desejada': 'Unidade'
    }

    cols_exist = [c for c in colunas_interesse.keys() if c in df.columns]
    if not cols_exist:
        raise ValueError('Nenhuma das colunas de interesse foi encontrada na base.')

    df_filtrado = df[cols_exist].copy()
    df_renomeado = df_filtrado.rename(columns=colunas_interesse)

    # Convert datetime
    df_renomeado['Data_Criacao'] = pd.to_datetime(df_renomeado.get('Data_Criacao'), errors='coerce')
    if 'Data_Fechamento' in df_renomeado.columns:
        df_renomeado['Data_Fechamento'] = pd.to_datetime(df_renomeado.get('Data_Fechamento'), errors='coerce')

    # Enriquecimento
    df_renomeado['Unidade'] = df_renomeado.get('Unidade').fillna('Sem Unidade').astype(str).str.strip()
    df_renomeado['Fonte_Trafego'] = df_renomeado.get('Fonte_Trafego').fillna('Sem Fonte').astype(str).str.strip()
    df_renomeado['Etapa_Negocio'] = df_renomeado.get('Etapa_Negocio').fillna('Sem Etapa').astype(str).str.strip()

    df_renomeado['Ano'] = df_renomeado['Data_Criacao'].dt.year
    df_renomeado['Mes'] = df_renomeado['Data_Criacao'].dt.to_period('M').apply(lambda r: r.start_time)
    df_renomeado['Semana'] = (df_renomeado['Data_Criacao'] - pd.to_timedelta((df_renomeado['Data_Criacao'].dt.weekday + 1) % 7, unit='D')).dt.normalize()

    # Conversão heurística
    final_kw = ['won', 'fechado', 'convertido', 'cliente', 'closed won', 'closedwon', 'negócio ganho']
    df_renomeado['_is_converted'] = df_renomeado['Etapa_Negocio'].astype(str).str.lower().apply(lambda s: any(k in s for k in final_kw))

    # Filtra para unidades próprias (se lista fornecida)
    if unidades_proprias:
        unidades_norm = [u.upper().strip() for u in unidades_proprias]
        df_renomeado['Unidade_norm'] = df_renomeado['Unidade'].astype(str).str.upper().str.strip()
        before_n = len(df_renomeado)
        df_renomeado = df_renomeado[df_renomeado['Unidade_norm'].isin(unidades_norm)].copy()
        after_n = len(df_renomeado)
        print(f'Filtrado unidades próprias: antes={before_n}, depois={after_n}')
        df_renomeado.drop(columns=['Unidade_norm'], inplace=True, errors='ignore')

    # Prepara df_export com colunas finais
    cols_final = [
        'ID_Negocio','Nome_Negocio','Data_Criacao','Data_Fechamento',
        'Unidade','Fonte_Trafego','Fonte_Detalhe_1','Fonte_Detalhe_2',
        'Etapa_Negocio','Ano','Mes','Semana','_is_converted'
    ]
    cols_final_existing = [c for c in cols_final if c in df_renomeado.columns]
    df_export = df_renomeado[cols_final_existing].copy()

    # Validação e coerção de esquema para Looker
    def validate_and_coerce(df):
        report = []
        # Data_Criacao -> datetime
        if 'Data_Criacao' in df.columns:
            df['Data_Criacao'] = pd.to_datetime(df['Data_Criacao'], errors='coerce')
            n_null = df['Data_Criacao'].isna().sum()
            report.append(f"Data_Criacao coerced to datetime (nulos={n_null})")
        else:
            report.append("FALTA: Data_Criacao")

        # Data_Fechamento -> datetime (if present)
        if 'Data_Fechamento' in df.columns:
            df['Data_Fechamento'] = pd.to_datetime(df['Data_Fechamento'], errors='coerce')

        # ID_Negocio -> string
        if 'ID_Negocio' in df.columns:
            df['ID_Negocio'] = df['ID_Negocio'].astype(str)
        else:
            report.append("AVISO: ID_Negocio ausente; usar índice gerado")

        # _is_converted -> int 0/1
        if '_is_converted' in df.columns:
            df['_is_converted'] = df['_is_converted'].astype(int).clip(0,1)
        else:
            df['_is_converted'] = 0
            report.append("AVISO: _is_converted não encontrado; preenchido com 0")

        # Mes -> month period start (datetime)
        if 'Mes' in df.columns:
            try:
                df['Mes'] = pd.to_datetime(df['Mes'])
                # convert to first day of month
                df['Mes'] = df['Mes'].dt.to_period('M').apply(lambda r: r.start_time)
            except Exception:
                report.append('AVISO: falha ao coerzir Mes para início do mês')

        # Semana -> date normalized
        if 'Semana' in df.columns:
            df['Semana'] = pd.to_datetime(df['Semana'], errors='coerce').dt.normalize()

        # Unidade, Fonte_Trafego, Etapa_Negocio as strings
        for c in ['Unidade','Fonte_Trafego','Etapa_Negocio','Fonte_Detalhe_1','Fonte_Detalhe_2']:
            if c in df.columns:
                df[c] = df[c].astype(str)

        print('\nSchema validation report:')
        for line in report:
            print('-', line)
        return df

    df_export = validate_and_coerce(df_export)

    # Consolidados
    if 'Unidade' in df_export.columns:
        df_unit_summary = df_export.groupby('Unidade').agg(
            Leads=('ID_Negocio', 'count') if 'ID_Negocio' in df_export.columns else ('Data_Criacao','count'),
            Converted=('_is_converted', 'sum') if '_is_converted' in df_export.columns else ('Data_Criacao','count')
        ).reset_index()
        df_unit_summary['Conversion_Rate'] = df_unit_summary['Converted'] / df_unit_summary['Leads']
        totals = pd.DataFrame({
            'Unidade': ['TOTAL'],
            'Leads': [df_unit_summary['Leads'].sum()],
            'Converted': [df_unit_summary['Converted'].sum()]
        })
        totals['Conversion_Rate'] = totals['Converted'] / totals['Leads']
        df_consolidado = pd.concat([df_unit_summary, totals], ignore_index=True)
    else:
        df_consolidado = pd.DataFrame()

    agg_semanal = df_export.groupby('Semana').size().reset_index(name='Leads').sort_values('Semana') if 'Semana' in df_export.columns else pd.DataFrame()
    agg_mensal = df_export.groupby('Mes').size().reset_index(name='Leads').sort_values('Mes') if 'Mes' in df_export.columns else pd.DataFrame()
    agg_fonte = df_export.groupby('Fonte_Trafego').size().reset_index(name='Leads').sort_values('Leads', ascending=False) if 'Fonte_Trafego' in df_export.columns else pd.DataFrame()
    agg_etapa = df_export.groupby('Etapa_Negocio').size().reset_index(name='Leads').sort_values('Leads', ascending=False) if 'Etapa_Negocio' in df_export.columns else pd.DataFrame()

    total_leads = len(df_export)
    total_converted = int(df_export['_is_converted'].sum()) if '_is_converted' in df_export.columns else 0
    overall_conversion = total_converted / total_leads if total_leads>0 else 0
    visao_metrics = pd.DataFrame({
        'Metric': ['Total Leads', 'Total Converted', 'Overall Conversion Rate'],
        'Value': [total_leads, total_converted, overall_conversion]
    })
    top_fontes = agg_fonte.head(10) if not agg_fonte.empty else pd.DataFrame()

    # Escreve Excel com múltiplas abas
    print(f'Escrevendo Excel em: {output_path}')
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_export.to_excel(writer, sheet_name='base_all', index=False)
        if not df_consolidado.empty:
            df_consolidado.to_excel(writer, sheet_name='consolidado', index=False)
        visao_metrics.to_excel(writer, sheet_name='visao_geral', index=False, startrow=0)
        if not top_fontes.empty:
            top_fontes.to_excel(writer, sheet_name='visao_geral', index=False, startrow=len(visao_metrics)+3)
        if 'Unidade' in df_export.columns:
            for unit in sorted(df_export['Unidade'].dropna().unique()):
                sheet_name = str(unit)[:31].replace('/', '_').replace('\\', '_')
                df_unit = df_export[df_export['Unidade'] == unit].copy()
                if df_unit.empty:
                    continue
                df_unit.to_excel(writer, sheet_name=sheet_name, index=False)
        if not agg_semanal.empty:
            agg_semanal.to_excel(writer, sheet_name='consolidado_semanal', index=False)
        if not agg_mensal.empty:
            agg_mensal.to_excel(writer, sheet_name='consolidado_mensal', index=False)
        if not agg_fonte.empty:
            agg_fonte.to_excel(writer, sheet_name='consolidado_fonte', index=False)
        if not agg_etapa.empty:
            agg_etapa.to_excel(writer, sheet_name='consolidado_etapa', index=False)

    print('Export concluído com sucesso.')


if __name__ == '__main__':
    # Caminhos padrão
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_input = r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_ATUAL.csv"
    default_output = os.path.join(os.path.dirname(script_dir), 'outputs', 'hubspot_dashboard_base.xlsx')
    unidades_proprias = ['VILA LEOPOLDINA','MORUMBI','ITAIM BIBI','PACAEMBU','PINHEIROS','JARDINS','PERDIZES','SANTANA']
    os.makedirs(os.path.dirname(default_output), exist_ok=True)
    try:
        generate_looker_base(default_input, default_output, unidades_proprias=unidades_proprias)
    except FileNotFoundError:
        print(f"ERRO: arquivo não encontrado: {default_input}")
    except Exception as e:
        print(f"Erro ao gerar base: {e}")