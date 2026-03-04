import pandas as pd
import os
import re
from pathlib import Path
from datetime import datetime

def get_cycle_dates(cycle_identifier):
    """
    Retorna data de início e fim esperadas para um ciclo (ex: 25.1 -> Out/24 a Mar/25).
    Aceita formatos: '25.1', '25_1', '23_1', etc.
    """
    # Normaliza separadores e procura padrão
    s = str(cycle_identifier).replace('_', '.')
    match = re.search(r'(\d{2})\.1', s)
    if match:
        year_suffix = int(match.group(1))
        # Lógica: Ciclo 26.1 começa em Outubro de 2025 (2000 + 26 - 1)
        start_year = 2000 + year_suffix - 1
        start_date = pd.Timestamp(f'{start_year}-10-01')
        end_date = pd.Timestamp(f'{start_year + 1}-02-20') + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        return start_date, end_date
    return None, None

def gerar_excel_consolidado(matriculas_store, leads_store, errors_store, output_filename, cutoff_note="YTD ate 06/02 (todos os ciclos)"):
    print(f"\n   [INFO] Gerando Excel Consolidado: {output_filename}...")
    
    # Identifica todos os ciclos presentes
    # Normaliza chaves (remove sufixos se houver, embora agora estejamos usando chaves limpas)
    cycles_mat = set(matriculas_store.keys())
    cycles_lead = set(leads_store.keys())
    cycles_err = set(errors_store.keys())
    all_cycles_raw = sorted(list(cycles_mat | cycles_lead | cycles_err))
    # Exclui ciclos auxiliares (ex: 25.1_FULL) das abas principais
    reporting_cycles = [c for c in all_cycles_raw if not c.endswith('_FULL')]
    
    # --- 1. ABA RESUMO ---
    lead_col = f"Leads ({cutoff_note})"
    matr_col = f"Matriculas ({cutoff_note})"
    err_col = f"Erros (Fech < Criacao) ({cutoff_note})"
    cycle_stats = {}
    for cycle in reporting_cycles:
        df_l = leads_store.get(cycle, pd.DataFrame())
        total_leads = len(df_l)
        
        df_m = matriculas_store.get(cycle, pd.DataFrame())
        if 'Nome do negócio' in df_m.columns:
            df_m = df_m.drop_duplicates(subset=['Nome do negócio'])
        total_matr = len(df_m)
        
        df_e = errors_store.get(cycle, pd.DataFrame())
        total_errors = len(df_e)
        
        conv_rate = total_matr / total_leads if total_leads > 0 else 0
        error_rate = total_errors / total_matr if total_matr > 0 else 0
        cycle_stats[cycle] = {
            'Leads': total_leads,
            'Matriculas': total_matr,
            'Erros (Fech < Criação)': total_errors,
            '% Erros': error_rate,
            'Conversao': conv_rate
        }
    
    df_summary = pd.DataFrame(cycle_stats).T
    df_summary.index.name = 'Ciclo'
    df_summary = df_summary.reset_index()

    # Deixa explicito no cabecalho como a volumetria foi calculada
    df_summary = df_summary.rename(columns={
        'Leads': lead_col,
        'Matriculas': matr_col,
        'Erros (Fech < Criação)': err_col
    })
    
    # Adiciona Variação % no Resumo (Linha a Linha)
    df_summary['Cresc. Leads %'] = df_summary[lead_col].pct_change().fillna(0)
    df_summary['Cresc. Matrículas %'] = df_summary[matr_col].pct_change().fillna(0)
    df_summary['Cresc. Erros %'] = df_summary[err_col].pct_change().fillna(0)

    # (O restante do arquivo original foi copiado; mantido função principal e exportações)

def verificar_diretorio():
    """Verifica diretórios esperados e lista arquivos CSV encontrados.

    Retorna o Path do diretório encontrado ou None se nenhum for localizado.
    """
    project_root = Path(__file__).resolve().parents[2]

    # Constroi o `Path` absoluto apontando para o CSV centralizado
    # e verifica se o arquivo existe no sistema de arquivos.
    # - Se existir: registra informação e retorna o `Path` (para ser carregado
    #   ou processado posteriormente);
    # - Se não existir: segue para procurar nas pastas locais de fallback.
    # Prioriza o arquivo de negócios perdidos indicado pelo usuário, mantendo o
    # arquivo `hubspot_leads_atual.csv` como fallback.
    central_candidates = [
        Path(r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_negocios_perdidos_ATUAL.csv"),
        Path(r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_atual.csv"),
    ]

    for central_file in central_candidates:
        if central_file.exists():
            print(f"[INFO] Arquivo central encontrado: {central_file}")
            return central_file

    candidates = [
        project_root / 'data' / 'backup' / 'realizados',
        project_root / 'data' / 'realizados',
        project_root / 'data'
    ]

    for candidate in candidates:
        if candidate.exists() and any(candidate.rglob('*.csv')):
            csvs = sorted(str(p.relative_to(project_root)) for p in candidate.rglob('*.csv'))
            print(f"[INFO] Diretório encontrado: {candidate}")
            print(f"[INFO] Arquivos CSV encontrados ({len(csvs)}):")
            for f in csvs:
                print(f"  - {f}")
            return candidate

    print("[WARN] Nenhum diretório com arquivos CSV foi encontrado entre os candidatos:")
    for c in candidates:
        print(f"  - {c}")
    return None


if __name__ == "__main__":
    verificar_diretorio()
