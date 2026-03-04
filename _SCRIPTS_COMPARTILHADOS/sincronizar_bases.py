"""
Script de Sincronização Automática de Bases
Atualiza as pastas Data/ dos projetos com as versões mais recentes de _DADOS_CENTRALIZADOS
Execute este script sempre que atualizar uma base no repositório central
"""

import shutil
from pathlib import Path
from datetime import datetime

ROOT = Path(r"C:\Users\a483650\Projetos")
CENTRAL = ROOT / "_DADOS_CENTRALIZADOS"

# Mapeamento: Fonte → Destinos
SYNC_MAP = {
    # HubSpot Leads
    "hubspot/hubspot_leads_ATUAL.csv": [
        "01_Helio_ML_Producao/Data/hubspot_leads.csv",
        "02_Pipeline_Midia_Paga/data/hubspot_leads.csv",
        "03_Analises_Operacionais/eficiencia_canal/data/hubspot_leads.csv",
    ],
    # HubSpot Negócios Perdidos
    "hubspot/hubspot_negocios_perdidos_ATUAL.csv": [
        "01_Helio_ML_Producao/Data/hubspot_negocios_perdidos.csv",
    ],
    # Matrículas
    "matriculas/matriculas_finais_ATUAL.csv": [
        "01_Helio_ML_Producao/Data/matriculas_finais_limpo.csv",
    ],
    "matriculas/matriculas_finais_ATUAL.xlsx": [
        "01_Helio_ML_Producao/Data/matriculas_finais.xlsx",
    ],
    # Marketing
    "marketing/meta_ads_ATUAL.csv": [
        "02_Pipeline_Midia_Paga/data/meta_dataset.csv",
    ],
    "marketing/google_ads_ATUAL.csv": [
        "02_Pipeline_Midia_Paga/data/googleads_dataset.csv",
    ],
}

def sincronizar_bases():
    """Sincroniza todas as bases dos projetos com o repositório central"""
    print("="*80)
    print("SINCRONIZAÇÃO DE BASES DE DADOS")
    print("="*80)
    print(f"\nRepositório Central: {CENTRAL}")
    print(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
    
    sucesso = 0
    erros = 0
    
    for fonte, destinos in SYNC_MAP.items():
        arquivo_fonte = CENTRAL / fonte
        
        if not arquivo_fonte.exists():
            print(f"\n⚠️  AVISO: Fonte não encontrada: {fonte}")
            erros += 1
            continue
        
        print(f"\n📄 {fonte}")
        tamanho_mb = arquivo_fonte.stat().st_size / (1024**2)
        modificado = datetime.fromtimestamp(arquivo_fonte.stat().st_mtime)
        print(f"   Tamanho: {tamanho_mb:.2f} MB")
        print(f"   Modificado: {modificado.strftime('%d/%m/%Y %H:%M:%S')}")
        
        for destino_rel in destinos:
            destino = ROOT / destino_rel
            
            try:
                # Criar diretório se não existir
                destino.parent.mkdir(parents=True, exist_ok=True)
                
                # Copiar arquivo
                shutil.copy2(arquivo_fonte, destino)
                print(f"   ✅ Sincronizado: {destino_rel}")
                sucesso += 1
                
            except Exception as e:
                print(f"   ❌ Erro ao sincronizar {destino_rel}: {e}")
                erros += 1
    
    # Resumo
    print("\n" + "="*80)
    print("RESUMO DA SINCRONIZAÇÃO")
    print("="*80)
    print(f"✅ Sucesso: {sucesso} arquivos")
    print(f"❌ Erros: {erros} arquivos")
    
    if erros > 0:
        print("\n⚠️  Revise os erros acima e corrija os problemas")
    else:
        print("\n🎉 Todas as bases foram sincronizadas com sucesso!")
    
    return sucesso, erros

if __name__ == "__main__":
    sincronizar_bases()
