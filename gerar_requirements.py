"""
Script para gerar requirements.txt em cada projeto
Detecta automaticamente as dependências do .venv
"""

import os
import subprocess
from pathlib import Path

PROJETOS = [
    '01_Helio_ML_Producao',
    '02_Pipeline_Midia_Paga',
    '03_Analises_Operacionais',
    '06_Analise_Funil_RedBalloon',
]

VENV_NAMES = ['.venv', 'venv', 'env']

def gerar_requirements(projeto_path):
    """Gera requirements.txt a partir do .venv"""
    
    projeto = Path(projeto_path)
    
    # Procura venv
    venv_dir = None
    for name in VENV_NAMES:
        venv = projeto / name
        if venv.exists():
            venv_dir = venv
            break
    
    if not venv_dir:
        print(f"  ⚠️  Nenhum venv encontrado em {projeto.name}")
        return False
    
    # Determina o python correto baseado no OS
    if os.name == 'nt':  # Windows
        python_exe = venv_dir / 'Scripts' / 'python.exe'
    else:  # Linux/Mac
        python_exe = venv_dir / 'bin' / 'python'
    
    if not python_exe.exists():
        print(f"  ❌ Python executable não encontrado em {venv_dir}")
        return False
    
    # Gera requirements.txt
    req_file = projeto / 'requirements.txt'
    
    try:
        result = subprocess.run(
            [str(python_exe), '-m', 'pip', 'freeze'],
            capture_output=True,
            text=True,
            check=True
        )
        
        with open(req_file, 'w') as f:
            f.write(result.stdout)
        
        num_libs = len(result.stdout.strip().split('\n'))
        print(f"  ✅ requirements.txt gerado ({num_libs} dependências)")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"  ❌ Erro ao gerar requirements.txt:")
        print(f"     {e.stderr}")
        return False

def main():
    print("=" * 70)
    print("📦 GERADOR DE REQUIREMENTS.TXT")
    print("=" * 70)
    print()
    
    root = Path('.')
    sucesso = 0
    falha = 0
    
    for projeto in PROJETOS:
        projeto_path = root / projeto
        
        if not projeto_path.exists():
            print(f"⏭️  {projeto} - NÃO ENCONTRADO")
            continue
        
        print(f"📁 {projeto}")
        if gerar_requirements(projeto_path):
            sucesso += 1
        else:
            falha += 1
        print()
    
    print("=" * 70)
    print(f"✅ Gerados com sucesso: {sucesso}")
    if falha > 0:
        print(f"❌ Falharam: {falha}")
    print("=" * 70)

if __name__ == '__main__':
    main()
