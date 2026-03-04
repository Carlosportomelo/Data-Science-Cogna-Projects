import sys
import os
import pandas as pd
import numpy as np
import glob

# Adiciona o diretório atual ao path para importar saneamento_leads
sys.path.append(os.path.dirname(__file__))
from saneamento_leads import try_read_csv_with_fallback, BASE_DIR, META_DIR, parse_age_or_dob, read_meta_all

def analyze_meta_files():
    print(f"🚀 INICIANDO ANÁLISE DE CONTEÚDO (DATA vs IDADE)...")
    print(f"📂 Pasta Meta: {META_DIR}")
    
    files = glob.glob(os.path.join(META_DIR, '**', 'nbfo*'), recursive=True)
    files = [f for f in files if os.path.isfile(f)]
    
    if not files:
        print("❌ Nenhum arquivo 'nbfo*' encontrado.")
        return

    report = []

    for file_path in files:
        filename = os.path.basename(file_path)
        df = try_read_csv_with_fallback(file_path)
        
        if df is None:
            report.append({'Arquivo': filename, 'Status': 'Erro leitura'})
            continue

        # Normaliza colunas
        df.columns = [str(c).lower().strip() for c in df.columns]
        
        # Procura colunas candidatas
        col_nasc = next((c for c in df.columns if any(k in c for k in ['nasc', 'birth', 'data'])), None)
        col_idade = next((c for c in df.columns if any(k in c for k in ['idade', 'age']) and 'nasc' not in c), None)
        
        total = len(df)
        n_dates = 0
        n_ages = 0
        
        # 1. Analisa coluna de 'Nascimento' (pode conter data ou idade mascarada)
        if col_nasc:
            # parsed retorna tupla (datetime, age_int)
            parsed = df[col_nasc].apply(parse_age_or_dob)
            
            # Conta quantos são datas válidas
            n_dates = parsed.apply(lambda x: 1 if x[0] is not None else 0).sum()
            
            # Conta quantos são apenas números (idades) na coluna de data
            n_ages_in_nasc = parsed.apply(lambda x: 1 if x[0] is None and x[1] is not None else 0).sum()
        else:
            n_ages_in_nasc = 0

        # 2. Analisa coluna explícita de 'Idade'
        n_ages_explicit = 0
        if col_idade:
            def is_valid_age(x):
                try:
                    # Tenta pegar numero inteiro
                    s = str(x).strip()
                    if s.isdigit():
                        v = int(s)
                        return 1 if 0 < v < 120 else 0
                    return 0
                except:
                    return 0
            n_ages_explicit = df[col_idade].apply(is_valid_age).sum()

        total_ages = n_ages_in_nasc + n_ages_explicit
        
        # Determina o cenário do arquivo
        if n_dates > 0 and total_ages == 0:
            conclusion = "Somente Data"
        elif n_dates == 0 and total_ages > 0:
            conclusion = "Somente Idade"
        elif n_dates > 0 and total_ages > 0:
            conclusion = "Misto (Data e Idade)"
        elif n_dates == 0 and total_ages == 0:
            conclusion = "Vazio / Sem Info"
        else:
            conclusion = "Indeterminado"

        report.append({
            'Arquivo': filename,
            'Col. Nasc': col_nasc or '-',
            'Qtd Datas': n_dates,
            'Qtd Idades': total_ages,
            'Conclusão': conclusion
        })

    # Exibe relatório
    df_report = pd.DataFrame(report)
    print("\n📊 RELATÓRIO DETALHADO POR ARQUIVO:")
    pd.set_option('display.max_rows', None)
    pd.set_option('display.width', 1000)
    print(df_report[['Arquivo', 'Col. Nasc', 'Qtd Datas', 'Qtd Idades', 'Conclusão']].to_string(index=False))
    
    print("\n📢 VEREDITO FINAL:")
    if (df_report['Qtd Datas'] == 0).all():
        print("✅ CONFIRMADO: Nenhum arquivo possui Datas de Nascimento válidas. Todos contêm apenas Idade ou estão vazios.")
    else:
        print("❌ NEGATIVO: Alguns arquivos contêm Datas de Nascimento reais (veja a coluna 'Qtd Datas' acima).")

    # Análise Global Consolidada
    print("\n" + "="*60)
    print("🚀 ANÁLISE CONSOLIDADA (Todos os CSVs - Deduplicado por E-mail)")
    print("="*60)
    
    try:
        # Usa a função oficial de leitura para garantir o mesmo tratamento do script principal
        df_all = read_meta_all(META_DIR)
        
        has_dob = df_all['data_nascimento_meta'].notna()
        has_age = df_all['idade_meta'].notna() & (df_all['idade_meta'] != '')
        
        print(f"\n📊 TOTAL DE LEADS ÚNICOS ENCONTRADOS: {len(df_all)}")
        print(f"1. Leads com Data de Nascimento preenchida: {has_dob.sum()}")
        print(f"2. Leads com Somente Idade (sem Data): {(has_age & ~has_dob).sum()}")
        print(f"3. Leads com Ambos (Data + Idade): {(has_dob & has_age).sum()}")
        print(f"4. Leads sem nenhuma informação: {(~has_dob & ~has_age).sum()}")
        
    except Exception as e:
        print(f"❌ Erro ao gerar análise consolidada: {e}")

if __name__ == '__main__':
    analyze_meta_files()