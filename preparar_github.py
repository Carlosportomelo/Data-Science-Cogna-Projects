"""
Script para preparar projetos para GitHub - Remove dados sensíveis
⚠️  LGPD Complaint - Remove apenas dados, mantém código
"""

import os
import shutil
import sys
from pathlib import Path

# Pastas/arquivos a DELETAR (contêm dados sensíveis)
PASTAS_DELETAR = [
    'Data',
    'data',
    'Outputs',
    'outputs',
    'Output',
    'output',
    '_DADOS_CENTRALIZADOS',
    '_BACKUP_SEGURANCA',
    'datasets',
    'raw_data',
    '.claude',
]

# Extensões de arquivo a DELETAR (dados)
EXTENSOES_DELETAR = [
    '.csv',
    '.xlsx',
    '.xls',
    '.parquet',
    '.json',
    '.db',
    '.sqlite',
    '.mdb',
]

# Pastas que podem ter dados mas vamos manter se forem testes
PASTAS_REVISAR = [
    'scripts/data_samples',
    'examples',
    'test_data',
]

PROJETOS = [
    '01_Helio_ML_Producao',
    '02_Pipeline_Midia_Paga',
    '03_Analises_Operacionais',
    '06_Analise_Funil_RedBalloon',
]

def limpar_pasta(pasta_root, simular=True):
    """Remove arquivos/pastas sensíveis"""
    
    deletados = []
    skipped = []
    
    pasta = Path(pasta_root)
    if not pasta.exists():
        print(f"❌ Pasta não encontrada: {pasta}")
        return
    
    # 1. Deletar pastas inteiras
    for subfolder in PASTAS_DELETAR:
        target = pasta / subfolder
        if target.exists():
            if simular:
                print(f"  [SIMULADO] Deletaria: {target}")
                deletados.append(str(target))
            else:
                shutil.rmtree(target)
                print(f"  ✓ Deletada: {target}")
                deletados.append(str(target))
    
    # 2. Deletar arquivos de dados
    for arquivo in pasta.rglob('*'):
        if arquivo.is_file():
            if arquivo.suffix.lower() in EXTENSOES_DELETAR:
                if simular:
                    print(f"  [SIMULADO] Deletaria: {arquivo.relative_to(pasta)}")
                    deletados.append(str(arquivo))
                else:
                    arquivo.unlink()
                    print(f"  ✓ Deletado: {arquivo.relative_to(pasta)}")
                    deletados.append(str(arquivo))
    
    return deletados, skipped

def verificar_estrutura(pasta_root):
    """Verifica se pasta tem estrutura esperada"""
    
    pasta = Path(pasta_root)
    reqs = ['README.md', 'scripts']  # Mínimo esperado
    
    presentes = []
    ausentes = []
    
    for req in reqs:
        if (pasta / req).exists():
            presentes.append(req)
        else:
            ausentes.append(req)
    
    return presentes, ausentes

def main():
    """Função principal"""
    
    root = Path('.')
    
    print("=" * 70)
    print("🔐 PREPARADOR PARA GITHUB - Remove dados sensíveis")
    print("=" * 70)
    print()
    
    modo_simular = True  # Sempre simula primeiro
    
    for projeto in PROJETOS:
        projeto_path = root / projeto
        
        if not projeto_path.exists():
            print(f"⏭️  {projeto} - NÃO ENCONTRADO (pulando)")
            print()
            continue
        
        print(f"📁 Processando: {projeto}")
        print("-" * 70)
        
        # Verifica estrutura
        presentes, ausentes = verificar_estrutura(projeto_path)
        if ausentes:
            print(f"  ⚠️  Aviso: Faltam arquivos ({', '.join(ausentes)})")
        
        # Simula limpeza
        deletados, _ = limpar_pasta(projeto_path, simular=True)
        
        if deletados:
            print(f"\n  📊 Encontrados {len(deletados)} arquivo(s)/pasta(s) para deletar")
        else:
            print(f"  ✅ Pasta limpa ou já sem dados sensíveis")
        
        print()
    
    # Pergunta se quer realmente deletar
    print("=" * 70)
    print("⚠️  MODO SIMULAÇÃO: Acima estão os arquivos que SERIAM deletados")
    print()
    resposta = input("Deseja REALMENTE deletar esses arquivos/pastas? (s/n): ").lower()
    
    if resposta == 's':
        print("\n🔴 Deletando arquivos sensíveis...")
        print("=" * 70)
        
        total_deletados = 0
        for projeto in PROJETOS:
            projeto_path = root / projeto
            if projeto_path.exists():
                deletados, _ = limpar_pasta(projeto_path, simular=False)
                total_deletados += len(deletados)
        
        print()
        print(f"✅ {total_deletados} arquivo(s)/pasta(s) deletado(s) com sucesso!")
    else:
        print("\n❌ Operação cancelada. Nenhum arquivo foi deletado.")
    
    print()
    print("=" * 70)
    print("✨ Próximas etapas:")
    print("  1. Verifique os arquivos manualmente se necessário")
    print("  2. Crie requirements.txt em cada projeto (pip freeze > requirements.txt)")
    print("  3. git add .")
    print("  4. git commit -m 'Initial commit - removido dados sensíveis'")
    print("  5. git push origin main")
    print("=" * 70)

if __name__ == '__main__':
    main()
