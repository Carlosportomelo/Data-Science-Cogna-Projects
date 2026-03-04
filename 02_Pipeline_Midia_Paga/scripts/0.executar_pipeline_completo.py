

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
🚀 PIPELINE COMPLETO DE ANÁLISE DE MÍDIA PAGA
==============================================

Executa todos os scripts do pipeline em sequência:
1. Análise de Performance Meta Ads
2. Análise de Performance Google Ads  
3. Análise de Performance HubSpot (Integração)
4. Geração de Cenários Preditivos

Uso:
    python 0.executar_pipeline_completo.py

Autor: Sistema de Data Science
Data: Fevereiro 2026
"""

import subprocess
import sys
from pathlib import Path
from datetime import datetime

def print_header(texto, etapa=None):
    """Imprime cabeçalho formatado"""
    print("\n" + "="*80)
    if etapa:
        print(f"   {etapa} - {texto}")
    else:
        print(f"   {texto}")
    print("="*80 + "\n")

def executar_script(caminho_script, nome_etapa, numero_etapa, total_etapas):
    """
    Executa um script Python e verifica se teve sucesso
    
    Args:
        caminho_script: Path do script a executar
        nome_etapa: Nome descritivo da etapa
        numero_etapa: Número da etapa atual
        total_etapas: Total de etapas do pipeline
    
    Returns:
        bool: True se sucesso, False se erro
    """
    print_header(nome_etapa, f"{numero_etapa}/{total_etapas}")
    
    try:
        # Executa o script e captura a saída
        resultado = subprocess.run(
            [sys.executable, str(caminho_script)],
            cwd=caminho_script.parent.parent,  # Executa do diretório raiz do projeto
            capture_output=False,  # Mostra output em tempo real
            text=True
        )
        
        if resultado.returncode == 0:
            print(f"\n✅ Etapa {numero_etapa} concluída com sucesso!\n")
            return True
        else:
            print(f"\n❌ ERRO na Etapa {numero_etapa} - {nome_etapa}")
            print(f"   Código de saída: {resultado.returncode}")
            return False
            
    except Exception as e:
        print(f"\n❌ ERRO ao executar {nome_etapa}: {e}")
        return False

def main():
    """Executa o pipeline completo"""
    inicio = datetime.now()
    
    # Diretório base
    try:
        BASE_DIR = Path(__file__).parent.parent
    except NameError:
        BASE_DIR = Path.cwd()
    
    SCRIPTS_DIR = BASE_DIR / "scripts"
    
    # Banner inicial
    print("\n" + "="*80)
    print("   🚀 PIPELINE DE MÍDIA PAGA - EXECUÇÃO COMPLETA")
    print("="*80)
    print(f"\n📂 Diretório: {BASE_DIR}")
    print(f"🕐 Início: {inicio.strftime('%d/%m/%Y %H:%M:%S')}\n")
    
    # Definir etapas do pipeline
    etapas = [
        {
            "script": SCRIPTS_DIR / "1.analise_performance_meta.py",
            "nome": "Análise de Performance Meta Ads",
            "descricao": "Processa dados de campanhas do Facebook/Instagram"
        },
        {
            "script": SCRIPTS_DIR / "2.analise_performance_google.py",
            "nome": "Análise de Performance Google Ads",
            "descricao": "Processa dados de campanhas do Google"
        },
        {
            "script": SCRIPTS_DIR / "3.analise_performance_hubspot_FINAL_ID.py",
            "nome": "Análise de Performance HubSpot (Integração)",
            "descricao": "Integra dados de leads e atribui investimentos"
        },
        {
            "script": SCRIPTS_DIR / "4.gerar_cenarios_preditivos.py",
            "nome": "Geração de Cenários Preditivos",
            "descricao": "Gera projeções e análises preditivas"
        }
    ]
    
    # Verificar se todos os scripts existem
    print("🔍 Verificando scripts...\n")
    scripts_faltando = []
    for etapa in etapas:
        if not etapa["script"].exists():
            scripts_faltando.append(etapa["script"].name)
            print(f"   ❌ {etapa['script'].name} - NÃO ENCONTRADO")
        else:
            print(f"   ✅ {etapa['script'].name}")
    
    if scripts_faltando:
        print(f"\n❌ ERRO: {len(scripts_faltando)} script(s) não encontrado(s)")
        print("   Corrija os problemas antes de continuar.")
        return 1
    
    print("\n✅ Todos os scripts encontrados!\n")
    input("Pressione ENTER para iniciar o pipeline...")
    
    # Executar cada etapa
    total_etapas = len(etapas)
    etapas_sucesso = 0
    
    for i, etapa in enumerate(etapas, 1):
        sucesso = executar_script(
            etapa["script"],
            etapa["nome"],
            i,
            total_etapas
        )
        
        if not sucesso:
            print("\n" + "="*80)
            print(f"   ⚠️  PIPELINE INTERROMPIDO NA ETAPA {i}/{total_etapas}")
            print("="*80)
            print(f"\n❌ Erro em: {etapa['nome']}")
            print(f"📄 Script: {etapa['script'].name}\n")
            print("💡 Dicas:")
            print("   1. Verifique se as bases de dados estão atualizadas")
            print("   2. Confira se há erros no log acima")
            print("   3. Execute o script individualmente para mais detalhes")
            return 1
        
        etapas_sucesso += 1
    
    # Conclusão
    fim = datetime.now()
    duracao = fim - inicio
    
    print("\n" + "="*80)
    print("   🎉 PIPELINE EXECUTADO COM SUCESSO!")
    print("="*80)
    
    print(f"\n📊 Resumo da Execução:")
    print(f"   ✅ Etapas concluídas: {etapas_sucesso}/{total_etapas}")
    print(f"   🕐 Início: {inicio.strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"   🕐 Fim: {fim.strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"   ⏱️  Duração: {duracao.total_seconds():.1f} segundos")
    
    print(f"\n📁 Resultados gerados em: outputs/")
    print("\n📄 Arquivos principais criados:")
    print("   • meta_dataset_dashboard.xlsx")
    print("   • google_dashboard.xlsx")
    print("   • meta_googleads_blend_[data].xlsx")
    print("   • cenarios_preditivos_[data].xlsx")
    
    print("\n" + "="*80 + "\n")
    
    return 0

if __name__ == "__main__":
    try:
        codigo_saida = main()
        input("\nPressione ENTER para fechar...")
        sys.exit(codigo_saida)
    except KeyboardInterrupt:
        print("\n\n⚠️  Pipeline cancelado pelo usuário")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n❌ ERRO FATAL: {e}")
        import traceback
        traceback.print_exc()
        input("\nPressione ENTER para fechar...")
        sys.exit(1)
