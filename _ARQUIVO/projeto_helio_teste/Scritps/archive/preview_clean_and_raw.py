import pandas as pd
import glob
import os

# find raw file
raw = os.path.join('Data', 'hubspot_negocios_perdidos.csv')
if not os.path.exists(raw):
    print('Arquivo raw não encontrado:', raw)
else:
    df_raw = pd.read_csv(raw, dtype=str, low_memory=False)
    print('--- BASE BRUTA (hubspot_negocios_perdidos.csv) ---')
    print('Linhas:', len(df_raw))
    print('Colunas:', df_raw.shape[1])
    print('\nColunas:', ', '.join(list(df_raw.columns[:20])) + (', ...' if df_raw.shape[1]>20 else ''))
    print('\nAmostra (10 linhas):')
    print(df_raw.head(10).to_string(index=False))

# find latest sanitized file in Outputs/archive
san_files = glob.glob(os.path.join('Outputs', 'archive', 'hubspot_negocios_perdidos_sanitizado_*.csv'))
if not san_files:
    print('\nArquivo sanitizado não encontrado em Outputs/archive')
else:
    san_files.sort(key=os.path.getmtime, reverse=True)
    san = san_files[0]
    df_san = pd.read_csv(san, dtype=str, low_memory=False)
    print('\n--- BASE SANITIZADA (' + os.path.basename(san) + ') ---')
    print('Linhas:', len(df_san))
    print('Colunas:', df_san.shape[1])
    print('\nColunas:', ', '.join(list(df_san.columns[:20])) + (', ...' if df_san.shape[1]>20 else ''))
    print('\nAmostra (10 linhas):')
    print(df_san.head(10).to_string(index=False))
