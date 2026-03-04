import os
import glob
import pandas as pd
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, 'Outputs')
os.makedirs(OUT, exist_ok=True)

DATE = datetime.now().strftime('%Y-%m-%d')
OUT_XLSX = os.path.join(OUT, f'Operational_Pack_Executive_{DATE}.xlsx')

def latest(pattern):
    files = glob.glob(os.path.join(OUT, pattern))
    return max(files, key=os.path.getmtime) if files else None

def create_excel():
    scored_f = latest('executive_ml_scored_*.csv')
    top_f = latest('top_recommendations_*.csv')
    ml_base = latest('ml_training_base_*.csv')

    if not scored_f or not top_f:
        raise SystemExit('Scored or top recommendations CSV not found in Outputs/. Run ML pipeline first.')

    df_scored = pd.read_csv(scored_f, encoding='utf-8')
    df_top = pd.read_csv(top_f, encoding='utf-8')

    # Scored slim: pick essential columns
    essential = []
    for c in ['Record ID', 'Nome do negócio', 'Data de criação', 'Etapa do negócio', 'Fonte original do tráfego', 'Detalhamento da fonte original do tráfego 1', 'Unidade Desejada', 'motivo_perda', 'label', 'ML_Score']:
        if c in df_scored.columns:
            essential.append(c)
    df_scored_slim = df_scored[essential].copy()

    # Channel summary
    chan_col = 'Fonte original do tráfego' if 'Fonte original do tráfego' in df_scored.columns else None
    if chan_col:
        chan_summary = df_scored.groupby(chan_col).agg(
            Leads=('ML_Score', 'count'),
            Avg_Score=('ML_Score', 'mean'),
            Labeled=('label', lambda s: s.notna().sum())
        ).reset_index().sort_values('Leads', ascending=False)
    else:
        chan_summary = pd.DataFrame()

    # Create Excel
    with pd.ExcelWriter(OUT_XLSX, engine='openpyxl') as w:
        # Top recommendations (already top-n)
        df_top.to_excel(w, sheet_name='Top_Recommendations_Top250', index=False)
        # Scored slim
        df_scored_slim.to_excel(w, sheet_name='Scored_Leads_Slim', index=False)
        # Channel summary
        if not chan_summary.empty:
            chan_summary.to_excel(w, sheet_name='Channel_Summary', index=False)
        # Add a sheet with short instructions
        instr = pd.DataFrame({
            'Item': ['What is this file', 'Top sheet', 'Scored sheet', 'Channel summary', 'How to use'],
            'Description': [
                'Operational pack for units: top leads and scored list to prioritize contact',
                'Top 250 recommended leads (by ML score) — action immediately',
                'All scored leads with essential fields (use filter by score and unit)',
                'Summary by channel: counts and average ML score',
                'Assign top leads to teams, SLA < 2 hours, monitor conversion and retrain weekly'
            ]
        })
        instr.to_excel(w, sheet_name='README', index=False)

    print('Operational Excel created:', OUT_XLSX)
    return OUT_XLSX

if __name__ == '__main__':
    create_excel()
