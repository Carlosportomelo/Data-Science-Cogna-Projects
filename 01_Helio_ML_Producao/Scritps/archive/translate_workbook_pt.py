import pandas as pd
from datetime import datetime

src = 'Outputs/Executive_Package_AllSheets_2025-12-11.xlsx'
dst = 'Outputs/Pacote_Executivo_TodasAbas_2025-12-11_pt.xlsx'

sheet_name_map = {
    'Top50': 'Top50',
    'Top250': 'Top250',
    'Top1000': 'Top1000',
    'Top10_per_unit': 'Top10_por_unidade',
    'Qualification_All': 'Qualificacao_Todos',
    'Daily_Plan': 'Plano_Diario',
    'Weekly_Plan': 'Plano_Semanal',
    'Looker_Dataset': 'Dataset_Looker',
    'Top_Recommendations_Top250': 'Recomendacoes_Top250',
    'Scored_Leads_Slim': 'Leads_Pontuados_Slim',
    'Channel_Summary': 'Resumo_Canais',
    'README': 'README'
}

# header translations: only apply when column exists
col_map = {
    'Record ID': 'ID do Registro',
    'RecordID': 'ID do Registro',
    'record id': 'ID do Registro',
    'Lead ID': 'ID do Lead',
    'lead_id': 'ID do Lead',
    'Name': 'Nome do negócio',
    'name': 'Nome do negócio',
    'Stage': 'Etapa do negócio',
    'stage': 'Etapa do negócio',
    'Email': 'Email',
    'email': 'Email',
    'Phone': 'Telefone',
    'Telephone': 'Telefone',
    'rank': 'Posição',
    'Rank': 'Posição',
    'Owner': 'Proprietário do negócio',
    'owner': 'Proprietário do negócio',
    'Pipeline': 'Pipeline',
    'Value': 'Valor na moeda da empresa',
    'Created Date': 'Data de criação',
    'Created': 'Data de criação',
    'Data de criação': 'Data de criação',
    'ML_Score_0_100': 'ML_Score_0_100',
    'ML_Prob_Calibrated': 'ML_Prob_Calibrated',
}

try:
    wb = pd.read_excel(src, sheet_name=None)
except Exception as e:
    print('Erro ao abrir o arquivo origem:', e)
    raise

out_sheets = {}
for name, df in wb.items():
    new_name = sheet_name_map.get(name, name)
    # standardize columns
    new_columns = {}
    for c in df.columns:
        if c in col_map:
            new_columns[c] = col_map[c]
        elif isinstance(c, str) and c.strip() in col_map:
            new_columns[c] = col_map[c.strip()]
        else:
            new_columns[c] = c
    df = df.rename(columns=new_columns)
    out_sheets[new_name] = df

# save translated workbook
with pd.ExcelWriter(dst, engine='openpyxl') as w:
    for sheet_name, df in out_sheets.items():
        # sanitize sheet name length
        sname = sheet_name[:31]
        try:
            df.to_excel(w, sheet_name=sname, index=False)
        except Exception as e:
            df2 = df.astype(str)
            df2.to_excel(w, sheet_name=sname, index=False)

print('Arquivo traduzido salvo em:', dst)
