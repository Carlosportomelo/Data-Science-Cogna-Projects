import argparse
from pathlib import Path
import sys
import pandas as pd


def detect_column(df, keywords):
    cols = list(df.columns)
    lowcols = [c.lower() for c in cols]
    for kw in keywords:
        for i, c in enumerate(lowcols):
            if kw in c:
                return cols[i]
    return None


def write_tables_to_sheet(writer, sheet_name, tables, startrow=0):
    row = startrow
    for title, df in tables:
        # write title
        tmp = pd.DataFrame([ [title] ])
        tmp.to_excel(writer, sheet_name=sheet_name, header=False, index=False, startrow=row)
        row += 1
        df.to_excel(writer, sheet_name=sheet_name, startrow=row, index=True)
        row += len(df) + 3


def compute_and_write(input_path, output_dir):
    input_path = Path(input_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    df = pd.read_csv(input_path, low_memory=False)

    # detect columns
    date_col = detect_column(df, ['create', 'date', 'created', 'data'])
    unit_col = detect_column(df, ['unidade', 'unit', 'branch', 'location'])
    source_col = detect_column(df, ['source', 'origem', 'channel'])
    stage_col = detect_column(df, ['stage', 'deal', 'status', 'etapa'])

    if date_col is None:
        raise RuntimeError('Não foi possível detectar a coluna de data de criação no arquivo.')

    # parse dates
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df = df.dropna(subset=[date_col]).copy()

    # normalize unit
    if unit_col is None:
        df['unit_detected'] = 'Sem Unidade'
        unit_col = 'unit_detected'
    else:
        df[unit_col] = df[unit_col].astype(str).str.strip()

    if source_col is None:
        df['source_detected'] = 'Sem Fonte'
        source_col = 'source_detected'
    else:
        df[source_col] = df[source_col].astype(str).str.strip()

    if stage_col is None:
        df['stage_detected'] = 'Sem Etapa'
        stage_col = 'stage_detected'
    else:
        df[stage_col] = df[stage_col].astype(str).str.strip()

    # helpers
    # calcular início da semana considerando semana iniciando no domingo (inclui finais de semana)
    # domingo será considerado início da semana
    df['week'] = (df[date_col] - pd.to_timedelta((df[date_col].dt.weekday + 1) % 7, unit='D')).dt.normalize()
    df['month'] = df[date_col].dt.to_period('M').apply(lambda r: r.start_time)

    # consolidated metrics
    total_leads = len(df)
    weekly = df.groupby('week').size().rename('leads').reset_index()
    monthly = df.groupby('month').size().rename('leads').reset_index()
    source_counts = df.groupby(source_col).size().rename('leads').reset_index()
    stage_counts = df.groupby(stage_col).size().rename('leads').reset_index()

    # conversion detection: define final stages heuristically
    final_kw = ['won', 'fechado', 'convertido', 'cliente', 'closed won', 'closedwon']
    df['_is_converted'] = df[stage_col].astype(str).str.lower().apply(lambda s: any(k in s for k in final_kw))

    conv_by_source = df.groupby(source_col)['_is_converted'].agg(['sum','count']).reset_index()
    conv_by_source['conversion_rate'] = conv_by_source['sum'] / conv_by_source['count']
    # Formatar como porcentagem para layout profissional
    conv_by_source['conversion_rate'] = conv_by_source['conversion_rate'].apply(lambda x: f"{x:.1%}" if pd.notnull(x) else "0.0%")

    # per unit metrics
    units = df[unit_col].fillna('Sem Unidade').unique().tolist()

    # write excel workbook with a sheet per unit and a summary
    excel_path = output_dir / 'hubspot_dashboard_base.xlsx'
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        # Inicio sheet with consolidated tables
        tables = [
            ('Total Leads', pd.DataFrame({'total': [total_leads]})),
            ('Leads por Semana', weekly.sort_values('week', ascending=False).set_index('week')),
            ('Leads por Mês', monthly.sort_values('month', ascending=False).set_index('month')),
            ('Fonte - Quantidade', source_counts.set_index(source_col)),
            ('Etapa - Quantidade', stage_counts.set_index(stage_col)),
            ('Taxa de Conversão por Fonte', conv_by_source.set_index(source_col)[['sum','count','conversion_rate']])
        ]
        write_tables_to_sheet(writer, 'Visão Geral', tables)

        # per-unit sheets
        for u in units:
            sub = df[df[unit_col] == u].copy()
            weekly_u = sub.groupby('week').size().rename('leads').reset_index()
            monthly_u = sub.groupby('month').size().rename('leads').reset_index()
            source_u = sub.groupby(source_col).size().rename('leads').reset_index()
            stage_u = sub.groupby(stage_col).size().rename('leads').reset_index()
            conv_u = sub.groupby(source_col)['_is_converted'].agg(['sum','count']).reset_index()
            conv_u['conversion_rate'] = conv_u['sum'] / conv_u['count']
            conv_u['conversion_rate'] = conv_u['conversion_rate'].apply(lambda x: f"{x:.1%}" if pd.notnull(x) else "0.0%")

            tables_u = [
                ('Resumo - Total Leads', pd.DataFrame({'total': [len(sub)]})),
                ('Leads por Semana', weekly_u.set_index('week')),
                ('Leads por Mês', monthly_u.set_index('month')),
                ('Fonte - Quantidade', source_u.set_index(source_col)),
                ('Etapa - Quantidade', stage_u.set_index(stage_col)),
                ('Taxa de Conversão por Fonte', conv_u.set_index(source_col)[['sum','count','conversion_rate']])
            ]
            safe = str(u)[:31]
            write_tables_to_sheet(writer, safe, tables_u)

    print('Relatórios gerados com sucesso em:', output_dir)


def main():
    parser = argparse.ArgumentParser(description='Gerar base para dashboard a partir de arquivo de leads (HubSpot)')
    parser.add_argument('--input', '-i', default=r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_ATUAL.csv", help='Caminho para o arquivo CSV de leads')
    parser.add_argument('--output', '-o', default='geracao_dashboard/outputs', help='Pasta de saída onde os relatórios serão salvos')
    args = parser.parse_args()

    try:
        compute_and_write(args.input, args.output)
    except Exception as e:
        print('Erro:', e, file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
