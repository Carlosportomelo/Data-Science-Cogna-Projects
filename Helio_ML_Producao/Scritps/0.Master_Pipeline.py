"""
==============================================================================
MASTER PIPELINE - ORQUESTRADOR DE EXECUÇÃO HELIO
==============================================================================
Script: 0.Master_Pipeline.py
Objetivo: Executar todos os scripts de produção na sequência correta de
          dependência de dados.
          1. Scoring (Gera o dado)
          2. Listas Operacionais (Consomem o dado)
          3. Relatórios Gerenciais (Consomem o dado)
          4. Histórico e Validação (Arquivam o dado)
          5. Limpeza (Organiza a casa)
Data: 2026-01-27
==============================================================================
"""

import subprocess
import sys
import os
import time
from datetime import datetime

# Configuração de Caminhos
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Definição da Sequência de Execução
# Formato: (Nome do Arquivo, Crítico?)
# Se Crítico=True e falhar, o pipeline para.
PIPELINE = [
    # --- ETAPA 1: CORE (Geração do Dado) ---
    ("1.ML_Lead_Scoring.py", True),

    # --- ETAPA 2: OPERACIONAL (Listas para as Unidades) ---
    ("7.Gerar_Listas_Unidades.py", False),
    ("11.Gerar_Planilhas_Feedback.py", False),

    # --- ETAPA 3: GERENCIAL & ESTRATÉGICO ---
    ("4.Analise_Unidades.py", False),
    ("6.Relatorio_Executivo.py", False),
    ("Report_Gerencial.py", False),

    # --- ETAPA 4: HISTÓRICO & VALIDAÇÃO ---
    ("5.Consolidador_Historico.py", True),
    ("10.Validacao_Conversao_Helio.py", True),
    ("2.Blind_Test_Accuracy.py", False),

    # --- ETAPA 5: LIMPEZA FINAL ---
    ("00.Organizador_Limpeza.py", False)
]

def print_header(text):
    print("\n" + "="*80)
    print(f" {text}")
    print("="*80)

def run_script(script_name, is_critical):
    script_path = os.path.join(BASE_DIR, script_name)
    
    if not os.path.exists(script_path):
        print(f"❌ [ERRO] Arquivo não encontrado: {script_name}")
        if is_critical:
            return False
        return True # Se não é crítico, segue o baile

    print(f"🚀 Iniciando: {script_name}...")
    start_time = time.time()
    
    try:
        # Executa o script python como um subprocesso
        result = subprocess.run(
            [sys.executable, script_path],
            check=False, # Não crasha o master se o filho falhar, nós tratamos abaixo
            text=True
        )
        
        elapsed = time.time() - start_time
        
        if result.returncode == 0:
            print(f"✅ [SUCESSO] {script_name} finalizado em {elapsed:.2f}s")
            return True
        else:
            print(f"❌ [FALHA] {script_name} retornou erro (Código {result.returncode})")
            if is_critical:
                print("   ⛔ Erro crítico detectado. Interrompendo pipeline.")
                return False
            return True
            
    except Exception as e:
        print(f"❌ [EXCEPTION] Erro ao tentar executar {script_name}: {e}")
        if is_critical: return False
        return True

def main():
    print_header("HELIO GROWTH IA - INICIANDO PIPELINE DE EXECUÇÃO")
    print(f"Data: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Diretório: {BASE_DIR}")
    
    total_scripts = len(PIPELINE)
    sucessos = 0
    
    for i, (script, critical) in enumerate(PIPELINE, 1):
        print(f"\n--- Passo {i}/{total_scripts} ---")
        
        if run_script(script, critical):
            sucessos += 1
        else:
            print("\n⛔ PIPELINE ABORTADO DEVIDO A ERRO CRÍTICO.")
            sys.exit(1)
            
    print_header("RESUMO DA EXECUÇÃO")
    print(f"Scripts Executados: {total_scripts}")
    print(f"Sucessos: {sucessos}")
    
    if sucessos == total_scripts:
        print("\n✨ PIPELINE CONCLUÍDO COM SUCESSO TOTAL ✨")
    else:
        print("\n⚠️ PIPELINE CONCLUÍDO COM AVISOS (Alguns scripts não críticos falharam)")

if __name__ == "__main__":
    main()