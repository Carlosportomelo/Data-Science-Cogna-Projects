import pandas as pd
import os
from datetime import datetime
from pathlib import Path

def read_financial_totals(data_dir_path):
    """Lê os totais realizados das planilhas financeiras de Out a Dez."""
    metas_dir = Path(data_dir_path) / 'realizado_out_dez'
    meta_files = {
        10: '2025_10_FECHAMENTO_OUTUBRO_2025_2026 1.xlsx',
        11: '2025_11_FECHAMENTO_NOVEMBRO_2025_2026 1.xlsx',
        12: '2025_12_FECHAMENTO_DEZEMBRO_2025_2026 1.xlsx'
    }
    
    financial_realized = {}
    print(f"[INFO] Lendo dados financeiros de: {metas_dir}")
    
    for m_key, fname in meta_files.items():
        fpath = metas_dir / fname
        if fpath.exists():
            try:
                df_meta = pd.read_excel(fpath, sheet_name='RESUMO', header=None)
                # Mapeamento: Outubro (M=12), Nov/Dez (N=13)
                idx_real = 12 if m_key == 10 else 13
                # Pula cabeçalho e soma
                val = pd.to_numeric(df_meta.iloc[1:, idx_real], errors='coerce').sum()
                financial_realized[m_key] = int(val)
                print(f"   > Mês {m_key}: {int(val)} matrículas (Planilha)")
            except Exception as e:
                print(f"   [AVISO] Erro ao ler {fname}: {e}")
                financial_realized[m_key] = 0
    
    return financial_realized

def generate_audit_report(df, data_dir, output_file):
    print("\n" + "="*40)
    print("GERANDO RELATÓRIO DE AUDITORIA (SEPARADO)")
    print("="*40)
    
    # Filtro para Ciclo 26.1 (Out 25 - Jan 26)
    start_26 = pd.Timestamp('2025-10-01')
    
    # Data de corte dinâmica (Max do arquivo ou Hoje)
    max_date = df['Data de criação'].max()
    cutoff_26 = max_date if pd.notnull(max_date) else pd.Timestamp.now()
    # Trava no fim do ciclo
    end_26_cap = pd.Timestamp('2026-03-31') + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    if cutoff_26 > end_26_cap: cutoff_26 = end_26_cap
    
    def get_month_sort_key(date):
        m = date.month
        return m if m >= 10 else m + 12

    mask_enr = lambda d: d['Etapa do negócio'].astype(str).str.contains("MATRÍCULA CONCLUÍDA", case=False, na=False)
    
    df_26_real = df[(df['Data de criação'] >= start_26) & (df['Data de criação'] <= cutoff_26)].copy()
    df_26_real['Month_Sort'] = df_26_real['Data de criação'].apply(get_month_sort_key)
    
    # Realizado HubSpot por Mês (Baseado em Data de Fechamento)
    df_26_enr = df[mask_enr(df) & (df['Data de fechamento'] >= start_26) & (df['Data de fechamento'] <= cutoff_26)].copy()
    df_26_enr['Month_Sort'] = df_26_enr['Data de fechamento'].apply(get_month_sort_key)
    
    realizado_por_mes = df_26_enr.groupby('Month_Sort').size()
    
    metas_dir = Path(data_dir) / 'realizado_out_dez'
    meta_files = {
        10: '2025_10_FECHAMENTO_OUTUBRO_2025_2026 1.xlsx',
        11: '2025_11_FECHAMENTO_NOVEMBRO_2025_2026 1.xlsx',
        12: '2025_12_FECHAMENTO_DEZEMBRO_2025_2026 1.xlsx'
    }
    
    audit_rows = []
    
    for m_key, fname in meta_files.items():
        fpath = metas_dir / fname
        if fpath.exists():
            try:
                df_meta = pd.read_excel(fpath, sheet_name='RESUMO', header=None)
                
                def sum_col(idx):
                    if idx < df_meta.shape[1]:
                        return pd.to_numeric(df_meta.iloc[1:, idx], errors='coerce').sum()
                    return 0

                # Mapeamento de Colunas
                if m_key == 10: 
                    idx_total, idx_rema, idx_real, idx_meta = 7, 10, 12, 13
                else: 
                    idx_total, idx_rema, idx_real, idx_meta = 7, 10, 13, 14

                meta_total = sum_col(idx_total)
                meta_rema = sum_col(idx_rema)
                realizado_planilha = sum_col(idx_real)
                meta_novos = sum_col(idx_meta)
                
                realizado_hubspot = realizado_por_mes.get(m_key, 0)
                diff = realizado_planilha - realizado_hubspot
                
                audit_rows.append({
                    'Mês': {10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'}.get(m_key, 'Outros'),
                    'Realizado (Hubspot)': realizado_hubspot,
                    'Realizado (Planilha)': realizado_planilha,
                    'Diferença (Planilha - Hubspot)': diff,
                    '% Cobertura HubSpot': realizado_hubspot / realizado_planilha if realizado_planilha > 0 else 0,
                    '% Perda (Não Integrado)': diff / realizado_planilha if realizado_planilha > 0 else 0,
                    'Meta Novos (Planilha)': meta_novos,
                    'Meta Rema (Planilha)': meta_rema,
                    'Meta Total (Planilha)': meta_total
                })
            except Exception as e:
                print(f"[AVISO] Erro ao ler meta {fname}: {e}")

    df_audit = pd.DataFrame(audit_rows)
    
    if not df_audit.empty:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            wb = writer.book
            fmt_header = wb.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
            fmt_pct = wb.add_format({'num_format': '0.0%'})
            
            df_audit.to_excel(writer, sheet_name='Auditoria de Dados', index=False)
            ws = writer.sheets['Auditoria de Dados']
            for col, val in enumerate(df_audit.columns): ws.write(0, col, val, fmt_header)
            ws.set_column('A:A', 15); ws.set_column('B:D', 18); ws.set_column('E:F', 15, fmt_pct); ws.set_column('G:I', 18)
        print(f"Relatório de Auditoria gerado: {output_file}")

# Note: For brevity this file was copied from the original location and preserved.
