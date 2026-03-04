"""
Script de Validação Pós-Reorganização
Testa se os projetos principais conseguem acessar suas bases de dados
"""

import os
from pathlib import Path
import pandas as pd

ROOT = Path(r"C:\Users\a483650\Projetos")

def validar_arquivo(caminho, nome_exibicao):
    """Valida se arquivo existe e pode ser lido"""
    try:
        if not caminho.exists():
            print(f"   ❌ Não encontrado: {nome_exibicao}")
            return False
        
        tamanho_mb = caminho.stat().st_size / (1024**2)
        
        # Tentar ler primeiras linhas
        if caminho.suffix == '.csv':
            df = pd.read_csv(caminho, nrows=5, encoding='utf-8')
        elif caminho.suffix == '.xlsx':
            df = pd.read_excel(caminho, nrows=5)
        else:
            print(f"   ⚠️  Formato não testado: {nome_exibicao}")
            return True
        
        linhas = len(df)
        colunas = len(df.columns)
        
        print(f"   ✅ OK: {nome_exibicao} ({tamanho_mb:.2f} MB, {colunas} colunas)")
        return True
        
    except Exception as e:
        print(f"   ❌ Erro ao ler {nome_exibicao}: {str(e)[:80]}")
        return False

def validar_projeto_helio():
    """Valida Projeto Helio (01_Helio_ML_Producao)"""
    print("\n" + "="*80)
    print("📊 VALIDANDO: 01_Helio_ML_Producao (Produção)")
    print("="*80)
    
    base = ROOT / "01_Helio_ML_Producao" / "Data"
    
    arquivos = {
        "hubspot_leads.csv": base / "hubspot_leads.csv",
        "hubspot_negocios_perdidos.csv": base / "hubspot_negocios_perdidos.csv",
        "matriculas_finais_limpo.csv": base / "matriculas_finais_limpo.csv",
        "matriculas_finais.xlsx": base / "matriculas_finais.xlsx",
    }
    
    ok = 0
    erros = 0
    
    for nome, caminho in arquivos.items():
        if validar_arquivo(caminho, nome):
            ok += 1
        else:
            erros += 1
    
    print(f"\n   Resultado: {ok}/{len(arquivos)} arquivos OK")
    return erros == 0

def validar_pipeline_midia_paga():
    """Valida Pipeline Mídia Paga (02_Pipeline_Midia_Paga)"""
    print("\n" + "="*80)
    print("📊 VALIDANDO: 02_Pipeline_Midia_Paga (Produção)")
    print("="*80)
    
    base = ROOT / "02_Pipeline_Midia_Paga" / "data"
    
    arquivos = {
        "hubspot_leads.csv": base / "hubspot_leads.csv",
        "meta_dataset.csv": base / "meta_dataset.csv",
    }
    
    ok = 0
    erros = 0
    
    for nome, caminho in arquivos.items():
        if validar_arquivo(caminho, nome):
            ok += 1
        else:
            erros += 1
    
    print(f"\n   Resultado: {ok}/{len(arquivos)} arquivos OK")
    return erros == 0

def validar_dados_centralizados():
    """Valida repositório central"""
    print("\n" + "="*80)
    print("📊 VALIDANDO: _DADOS_CENTRALIZADOS (Repositório Central)")
    print("="*80)
    
    central = ROOT / "_DADOS_CENTRALIZADOS"
    
    arquivos = {
        "hubspot_leads_ATUAL.csv": central / "hubspot" / "hubspot_leads_ATUAL.csv",
        "hubspot_negocios_perdidos_ATUAL.csv": central / "hubspot" / "hubspot_negocios_perdidos_ATUAL.csv",
        "matriculas_finais_ATUAL.csv": central / "matriculas" / "matriculas_finais_ATUAL.csv",
        "matriculas_finais_ATUAL.xlsx": central / "matriculas" / "matriculas_finais_ATUAL.xlsx",
        "meta_ads_ATUAL.csv": central / "marketing" / "meta_ads_ATUAL.csv",
    }
    
    ok = 0
    erros = 0
    
    for nome, caminho in arquivos.items():
        if validar_arquivo(caminho, nome):
            ok += 1
        else:
            erros += 1
    
    print(f"\n   Resultado: {ok}/{len(arquivos)} arquivos OK")
    return erros == 0

def main():
    print("="*80)
    print("🔍 VALIDAÇÃO PÓS-REORGANIZAÇÃO")
    print("="*80)
    print(f"\nDiretório: {ROOT}")
    print("Testando acesso às bases de dados dos projetos principais\n")
    
    resultados = []
    
    # Validar repositório central
    resultados.append(("_DADOS_CENTRALIZADOS", validar_dados_centralizados()))
    
    # Validar projetos em produção
    resultados.append(("01_Helio_ML_Producao", validar_projeto_helio()))
    resultados.append(("02_Pipeline_Midia_Paga", validar_pipeline_midia_paga()))
    
    # Resumo final
    print("\n" + "="*80)
    print("📋 RESUMO DA VALIDAÇÃO")
    print("="*80)
    
    for nome, ok in resultados:
        status = "✅ OK" if ok else "❌ FALHOU"
        print(f"{status} - {nome}")
    
    total_ok = sum(1 for _, ok in resultados if ok)
    total = len(resultados)
    
    print(f"\n🎯 Total: {total_ok}/{total} projetos validados com sucesso")
    
    if total_ok == total:
        print("\n🎉 REORGANIZAÇÃO CONCLUÍDA COM SUCESSO!")
        print("   Todos os projetos podem acessar suas bases de dados.")
    else:
        print("\n⚠️  Alguns projetos apresentaram erros.")
        print("   Revise os problemas acima antes de usar os scripts.")

if __name__ == "__main__":
    main()
