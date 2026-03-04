import os
import pandas as pd
import sys
import glob

# Adiciona diretório ao path para importar utilitários
sys.path.append(os.path.dirname(__file__))
from saneamento_leads import find_hubspot_file, try_read_csv_with_fallback, BASE_DIR, DATA_DIR, META_DIR

def main():
    print("🚀 FERRAMENTA DE BUSCA E VERIFICAÇÃO (HubSpot vs Meta)")
    
    # 1. Identificar Arquivo HubSpot
    # Tenta achar o arquivo específico mencionado ou o mais recente
    hub_file = find_hubspot_file(DATA_DIR, preferred_name="Novos Negócios 04_02 08h53 2")
    if not hub_file:
        print("❌ Arquivo HubSpot não encontrado em Data/.")
        return
    print(f"📂 Arquivo HubSpot Alvo: {os.path.basename(hub_file)}")
    
    # Ler HubSpot
    try:
        if hub_file.endswith(('.xls', '.xlsx')):
            try:
                df_hs = pd.read_excel(hub_file)
            except:
                df_hs = pd.read_html(hub_file)[0]
        else:
            df_hs = try_read_csv_with_fallback(hub_file)
    except Exception as e:
        print(f"❌ Erro ao ler HubSpot: {e}")
        return

    # 2. Identificar Arquivos Meta
    meta_files = glob.glob(os.path.join(META_DIR, '**', 'nbfo*'), recursive=True)
    print(f"📂 Arquivos Meta encontrados: {len(meta_files)}")
    
    # 3. Obter termo de busca
    if len(sys.argv) > 1:
        termo = sys.argv[1]
    else:
        print("\nDigite o termo para buscar (pode ser Nome, E-mail, Telefone ou 'coluna:idade'):")
        termo = input("> ").strip()
        
    if not termo:
        print("Busca cancelada.")
        return

    print(f"\n🔎 Buscando por: '{termo}'...")
    
    # --- BUSCA NO HUBSPOT ---
    print(f"\n--- RESULTADOS HUBSPOT ---")
    if termo.lower().startswith('coluna:'):
        col_name = termo.split(':')[1]
        cols = [c for c in df_hs.columns if col_name.lower() in str(c).lower()]
        if cols:
            print(f"✅ Colunas encontradas: {cols}")
        else:
            print("❌ Coluna não encontrada.")
    else:
        # Busca textual em todo o dataframe
        mask = df_hs.astype(str).apply(lambda x: x.str.contains(termo, case=False, na=False)).any(axis=1)
        found = df_hs[mask]
        if not found.empty:
            print(f"✅ Encontrado {len(found)} registro(s):")
            print(found.head(3).to_string())

    # --- BUSCA NO META ---
    print(f"\n--- RESULTADOS META ---")
    found_any = False
    for f in meta_files:
        try:
            df = try_read_csv_with_fallback(f, verbose=False)
            if df is None: continue
            
            if termo.lower().startswith('coluna:'):
                col_name = termo.split(':')[1]
                cols = [c for c in df.columns if col_name.lower() in str(c).lower()]
                if cols:
                    print(f"✅ Em {os.path.basename(f)}: Colunas {cols}")
                    found_any = True
            else:
                mask = df.astype(str).apply(lambda x: x.str.contains(termo, case=False, na=False)).any(axis=1)
                found = df[mask]
                if not found.empty:
                    print(f"✅ Em {os.path.basename(f)}: {len(found)} registro(s)")
                    # Mostra colunas relevantes
                    cols_show = [c for c in df.columns if any(k in str(c).lower() for k in ['mail','nome','name','nasc','birth','idade','age'])]
                    if not cols_show: cols_show = df.columns[:5]
                    print(found[cols_show].head(2).to_string(index=False))
                    found_any = True
        except:
            pass

if __name__ == "__main__":
    main()
