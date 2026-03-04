"""Gera uma amostra de sucesso (5 exemplos) com BÔNUS SECRETO (Varredura de Data Bruta)."""
from saneamento_leads import find_hubspot_file, read_hubspot, read_meta_all, compute_age_from_dob, BASE_DIR, DATA_DIR, META_DIR
import pandas as pd
import os
import glob

OUTPUT_DIR = os.path.join(BASE_DIR, 'Output')
os.makedirs(OUTPUT_DIR, exist_ok=True)

def executar_bonus_secreto(sample_df):
    """
    🕵️‍♂️ BÔNUS SECRETO: Varre os arquivos brutos novamente apenas para os leads da amostra
    para garantir que a data de nascimento exata (birth_date) seja recuperada.
    """
    print("\n🕵️‍♂️ Rodando Bônus Secreto (Deep Dive nos arquivos brutos)...")
    
    # Dicionário para guardar o ouro: {email: data_bruta}
    secret_vault = {}
    target_emails = sample_df['E-mail'].str.lower().str.strip().tolist()
    
    # Pega todos os arquivos brutos
    raw_files = glob.glob(os.path.join(META_DIR, '**', 'nbfo*.csv'), recursive=True) + \
                glob.glob(os.path.join(META_DIR, '**', 'NB*.CSV'), recursive=True)
    
    for f in raw_files:
        try:
            # Leitura forçada (bruta)
            try:
                df = pd.read_csv(f, sep='\t', encoding='utf-16')
            except:
                try:
                    df = pd.read_csv(f, sep='\t', encoding='utf-8')
                except:
                    df = pd.read_csv(f)
            
            # Normaliza colunas
            df.columns = [c.lower().strip() for c in df.columns]
            
            # Acha colunas críticas
            col_email = next((c for c in df.columns if 'mail' in c), None)
            col_nasc = next((c for c in df.columns if 'birth' in c or 'nasc' in c), None) # Prioriza birth_date
            
            if col_email and col_nasc:
                # Filtra apenas os leads da amostra neste arquivo
                df['email_clean'] = df[col_email].astype(str).str.lower().str.strip()
                hits = df[df['email_clean'].isin(target_emails)]
                
                for _, row in hits.iterrows():
                    email = row['email_clean']
                    data = row[col_nasc]
                    # Só guarda se tiver data válida
                    if pd.notna(data) and str(data).strip() != '':
                        secret_vault[email] = data
                        
        except Exception:
            continue
            
    # Aplica o bônus na amostra
    raw_vals = sample_df['E-mail'].map(secret_vault)
    sample_df['Data de Nascimento'] = raw_vals.fillna(sample_df['Data de Nascimento'])
    return sample_df

def main(save: bool = True, n: int = 5, target_col: str = None, sample_method: str = 'random'):
    print("🔎 Localizando arquivo 'Novos Negócios'...")
    hub_file = find_hubspot_file(DATA_DIR)
    if hub_file is None:
        print("❌ Arquivo 'Novos Negócios' não encontrado em Data/.")
        return
    print(f"📄 HubSpot: {os.path.basename(hub_file)}")

    df_hs = read_hubspot(hub_file)
    df_meta = read_meta_all(META_DIR)

    # Identifica coluna alvo
    col_patterns = ['data de nascimento do aluno', 'data de nascimento', 'nascimento', 'data_nasc', 'diário', 'diario']
    target = None
    if target_col:
        cols = list(df_hs.columns)
        for c in cols:
            if target_col.lower() in str(c).lower():
                target = c
                break
        if target is None and target_col in cols:
            target = target_col
    else:
        for c in df_hs.columns:
            cname = str(c).lower()
            if any(p in cname for p in col_patterns):
                target = c
                break

    if target is None:
        print("⚠️ Coluna alvo não encontrada. Use --target-col.")
        return

    # Filtra leads vazios
    def is_original_empty(val):
        if pd.isna(val): return True
        s = str(val).strip().lower()
        return s == '' or 'nenhum' in s

    df_hs_filtered = df_hs[df_hs[target].apply(is_original_empty)].copy()
    if df_hs_filtered.empty:
        print("⚠️ Nada a fazer (nenhum lead vazio).")
        return
    print(f"✅ Leads vazios para processar: {len(df_hs_filtered)}")

    # Merge
    df_final = pd.merge(df_hs_filtered, df_meta, on='key_email', how='left')
    
    # Lógica de Idade/Encontrado
    from saneamento_leads import apply_exact_fill
    df_final, _ = apply_exact_fill(df_final, df_meta)
    
    if 'Status' not in df_final.columns:
        df_final['Status'] = df_final['data_nascimento_meta'].apply(lambda x: 'ENCONTRADO' if pd.notna(x) else 'NÃO ENCONTRADO')

    df_success = df_final[df_final['Status'] == 'ENCONTRADO'].copy()
    if df_success.empty:
        print("⚠️ Nenhum sucesso no cruzamento padrão.")
        return

    # Cálculos Finais
    def idade_final(row):
        try:
            if 'idade' in row and pd.notna(row['idade']) and row['idade'] != '': return int(row['idade'])
            if 'idade_meta' in row and pd.notna(row['idade_meta']) and row['idade_meta'] != '': return int(row['idade_meta'])
            if pd.notna(row['data_nascimento_meta']): return compute_age_from_dob(row['data_nascimento_meta'])
        except: return None
        return None

    df_success['Idade do Aluno'] = df_success.apply(idade_final, axis=1)
    df_success['Data de Nascimento'] = df_success['data_nascimento_meta'].apply(lambda x: pd.to_datetime(x).date().isoformat() if pd.notna(x) else '')

    # Prepara output base
    # Preserva colunas originais do HubSpot (removendo auxiliares do script)
    cols_originais = [c for c in df_hs.columns if c not in ['key_email', 'phone_hs', 'phone_hs_norm']]
    
    out = df_success.copy()
    # Trata colisão de nome 'E-mail' se existir na base original
    if 'E-mail' in out.columns and 'key_email' in out.columns:
        out.rename(columns={'E-mail': 'E-mail (Original)'}, inplace=True)
        cols_originais = [c if c != 'E-mail' else 'E-mail (Original)' for c in cols_originais]
        
    out.rename(columns={'key_email': 'E-mail'}, inplace=True)
    
    # Monta lista final: E-mail (chave), Idade, Data, [Originais...]
    cols_final = ['E-mail', 'Idade do Aluno', 'Data de Nascimento'] + cols_originais
    # Garante que colunas existem
    cols_final = [c for c in cols_final if c in out.columns]
    
    sample_source = out[cols_final]

    # Seleciona Amostra
    if sample_method == 'first':
        sample = sample_source.head(n).reset_index(drop=True)
    else:
        sample = sample_source.sample(n=min(n, len(sample_source))).reset_index(drop=True)

    # ==========================================
    # 🔥 AQUI ENTRA O BÔNUS SECRETO 🔥
    # ==========================================
    sample = executar_bonus_secreto(sample)

    # [CORREÇÃO] Recalcula a idade caso o Bônus Secreto tenha encontrado uma data nova
    # para preencher lacunas onde a idade estava vazia.
    def recalcular_idade(row):
        current_age = row.get('Idade do Aluno')
        if pd.notna(current_age) and str(current_age).strip() != '':
            return current_age
        
        # Tenta calcular da data (que pode ter sido preenchida pelo bônus)
        dob = row.get('Data de Nascimento')
        if pd.notna(dob) and str(dob).strip() != '':
            try:
                dt = pd.to_datetime(dob, dayfirst=True, errors='coerce')
                if pd.notna(dt):
                    return compute_age_from_dob(dt)
            except:
                pass
        return current_age

    sample['Idade do Aluno'] = sample.apply(recalcular_idade, axis=1)

    # Garante que as colunas solicitadas tenham a mesma informação da data encontrada
    if 'Data de Nascimento' in sample.columns:
        sample['Data de Nascimento do Aluno'] = sample['Data de Nascimento']
        sample['Data de Nascimento do Candidato'] = sample['Data de Nascimento']

    # Imprime no terminal
    print('\n✅ Amostra de Sucesso (Com Bônus Secreto):')
    print(sample[['E-mail', 'Idade do Aluno', 'Data de Nascimento']].to_string(index=False))

    # Salva
    if save:
        out_path = os.path.join(OUTPUT_DIR, 'amostra_prova_de_conceito.csv')
        sample.to_csv(out_path, index=False, sep=';', encoding='utf-8-sig')
        print(f"\n💾 Amostra salva em: {out_path}")

    # Gera versão Excel Diogo (Mantida igual ao seu original)
    try:
        excel_path = os.path.join(OUTPUT_DIR, 'amostra_prova_de_conceito_diogo.xlsx')
        print(f"\n📦 Gerando Excel Diogo: {os.path.basename(excel_path)}")
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            # Aba Diogo_Preenchido (Renomeia colunas para padrão esperado pelo validador)
            df_diogo = sample.copy()
            df_diogo.rename(columns={'Idade do Aluno': 'idade'}, inplace=True)
            df_diogo.to_excel(writer, sheet_name='Diogo_Preenchido', index=False)
            
            # Aba Amostra
            sample.to_excel(writer, sheet_name='Amostra', index=False)
        print("✅ Excel gerado com sucesso.")
    except Exception as e:
        print(f"⚠️ Falha ao gerar Excel Diogo: {e}")

if __name__ == '__main__':
    # (Mantido igual)
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--target-col', dest='target_col', default=None)
    parser.add_argument('--n', dest='n', type=int, default=5)
    parser.add_argument('--no-save', dest='save', action='store_false')
    parser.add_argument('--sample-method', dest='sample_method', choices=['random', 'first'], default='random')
    args = parser.parse_args()
    main(save=args.save, n=args.n, target_col=args.target_col, sample_method=args.sample_method)