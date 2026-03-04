"""
==============================================================================
ORGANIZADOR DE LIMPEZA (HOUSEKEEPING)
==============================================================================
Script: 00.Organizador_Limpeza.py
Objetivo: Varrer as pastas de Output, manter apenas os arquivos mais recentes
          (de hoje ou os últimos 3) e mover o restante para a pasta 'Backup'.
          Ignora pastas que já possuem lógica própria de backup se estiverem limpas.
Data: 2026-01-27
==============================================================================
"""

import os
import shutil
import glob
from datetime import datetime

# ==============================================================================
# CONFIGURAÇÃO
# ==============================================================================
CAMINHO_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA_OUTPUT = os.path.join(CAMINHO_BASE, 'Outputs')

# Pastas Alvo (Onde a bagunça costuma acumular)
PASTAS_ALVO = [
    'Relatorios_Executivos',    # Script 6
    'Controle_Feedback',        # Script 11
    'Relatorio_Gerencial',      # Script Report_Gerencial
    'Relatorios_Unidades',      # Script 4 (Reforço)
    'Relatorios_ML',            # Script 1 (Reforço)
    'Dados_Scored',             # Script 1 (Reforço)
    'Relatorios_Blind_Test'     # Script 10 e 5 (Novo)
]

print("="*80)
print("ORGANIZADOR DE ARQUIVOS - LIMPEZA E BACKUP")
print("="*80)

def organizar_diretorio(nome_subpasta):
    caminho_completo = os.path.join(PASTA_OUTPUT, nome_subpasta)
    
    if not os.path.exists(caminho_completo):
        print(f"⚠️ Pasta não encontrada (pulando): {nome_subpasta}")
        return

    print(f"📂 Processando: {nome_subpasta}")
    
    # Criar pasta de Backup se não existir
    pasta_backup = os.path.join(caminho_completo, 'Backup')
    os.makedirs(pasta_backup, exist_ok=True)
    
    # Listar arquivos na raiz da pasta (ignora subpastas)
    arquivos = [
        os.path.join(caminho_completo, f) 
        for f in os.listdir(caminho_completo) 
        if os.path.isfile(os.path.join(caminho_completo, f))
    ]
    
    # Se não tem arquivos, pula
    if not arquivos:
        print("   -> Pasta vazia ou apenas com subpastas.")
        print("-" * 40)
        return

    # Ordenar por data de modificação (Mais recente primeiro)
    arquivos.sort(key=os.path.getmtime, reverse=True)
    
    # Critérios de Manutenção
    data_hoje = datetime.now().date()
    contagem_movidos = 0
    contagem_mantidos = 0
    
    for i, arquivo in enumerate(arquivos):
        nome_arquivo = os.path.basename(arquivo)
        timestamp = os.path.getmtime(arquivo)
        data_arquivo = datetime.fromtimestamp(timestamp).date()
        
        # Regra: Manter se for de HOJE ou se for um dos 3 mais recentes
        if data_arquivo == data_hoje or i < 3:
            contagem_mantidos += 1
        else:
            # Mover para Backup
            destino = os.path.join(pasta_backup, nome_arquivo)
            try:
                if os.path.exists(destino): os.remove(destino) # Sobrescreve se já existir no backup
                shutil.move(arquivo, destino)
                contagem_movidos += 1
            except Exception as e:
                print(f"   ❌ Erro ao mover {nome_arquivo}: {e}")

    if contagem_movidos > 0:
        print(f"   ✅ Organizado: {contagem_movidos} arquivos movidos para Backup.")
        print(f"   ℹ️  Mantidos: {contagem_mantidos} arquivos recentes.")
    else:
        print("   ✨ Pasta já está organizada.")
    
    print("-" * 40)

def limpar_raiz():
    """Varre a raiz de Outputs e move arquivos soltos para suas pastas corretas."""
    print("🧹 Verificando arquivos soltos na raiz 'Outputs'...")
    
    # Mapeamento: Padrão do nome do arquivo -> Pasta de Destino
    regras_movimentacao = {
        'Blind_Test': 'Relatorios_Blind_Test',
        'Painel_Evolucao': 'DB_HELIO',
        'Evolucao_Temporal': 'DB_HELIO',
        'leads_scored': 'Dados_Scored',
        'Relatorio_Gerencial': 'Relatorio_Gerencial',
        'Relatorio_Executivo': 'Relatorios_Executivos',
        'Feedback': 'Controle_Feedback',
        'Relatorio_Unidades': 'Relatorios_Unidades'
    }
    
    arquivos_raiz = [f for f in os.listdir(PASTA_OUTPUT) if os.path.isfile(os.path.join(PASTA_OUTPUT, f))]
    
    movidos = 0
    for arquivo in arquivos_raiz:
        moved = False
        for padrao, pasta_destino in regras_movimentacao.items():
            if padrao.lower() in arquivo.lower():
                origem = os.path.join(PASTA_OUTPUT, arquivo)
                destino_dir = os.path.join(PASTA_OUTPUT, pasta_destino)
                destino_arquivo = os.path.join(destino_dir, arquivo)
                
                os.makedirs(destino_dir, exist_ok=True)
                
                try:
                    shutil.move(origem, destino_arquivo)
                    print(f"   -> Movido: {arquivo}  >>>  {pasta_destino}/")
                    movidos += 1
                    moved = True
                    break
                except Exception as e:
                    print(f"   ❌ Erro ao mover {arquivo}: {e}")
        
        if not moved and not arquivo.startswith('.'): # Ignora arquivos de sistema
            print(f"   ⚠️  Arquivo sem pasta definida (mantido na raiz): {arquivo}")

    if movidos == 0:
        print("   ✨ Raiz limpa. Nenhum arquivo solto encontrado.")
    print("-" * 40)

# Executar para todas as pastas alvo
limpar_raiz() # Executa primeiro a limpeza da raiz
for pasta in PASTAS_ALVO:
    organizar_diretorio(pasta)

print("\n[CONCLUÍDO] Todas as pastas foram verificadas.")