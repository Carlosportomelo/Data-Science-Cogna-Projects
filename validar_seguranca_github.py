"""
Script de validação pré-GitHub
Verifica se tudo está seguro para fazer push
"""

import os
from pathlib import Path

ARQUIVOS_PERIGOSOS = [
    '*.csv', '*.xlsx', '*.xls',
    '.env', '*.key', '*.pem',
    'credentials.json', 'config.local.py'
]

PASTAS_PERIGOSAS = [
    'data', 'Data', 'outputs', 'Outputs',
    '_DADOS_', '_BACKUP_', 'datasets'
]

PROJETOS = [
    '01_Helio_ML_Producao',
    '02_Pipeline_Midia_Paga',
    '03_Analises_Operacionais',
    '06_Analise_Funil_RedBalloon',
]

def verificar_seguranca(projeto_path):
    """Verifica se tem dados sensíveis"""
    
    problemas = []
    projeto = Path(projeto_path)
    
    # Verifica pastas perigosas
    for pattern in PASTAS_PERIGOSAS:
        for pasta in projeto.glob(f'*{pattern}*'):
            if pasta.is_dir():
                # Conta arquivos dentro
                arquivos = list(pasta.rglob('*'))
                problemas.append({
                    'tipo': 'pasta_perigosa',
                    'path': str(pasta.relative_to(projeto)),
                    'arquivos': len([a for a in arquivos if a.is_file()])
                })
    
    # Verifica arquivos perigosos (extensões sensíveis)
    extensoes_sensíveis = ['.csv', '.xlsx', '.xls', '.parquet', '.db']
    for ext in extensoes_sensíveis:
        for arquivo in projeto.rglob(f'*{ext}'):
            if arquivo.is_file():
                # Ignora pastas específicas que podem ter exemplos teste
                path_str = str(arquivo.relative_to(projeto)).lower()
                if 'example' not in path_str and '_example' not in path_str:
                    problemas.append({
                        'tipo': 'arquivo_sensivel',
                        'path': str(arquivo.relative_to(projeto)),
                        'tamanho_kb': arquivo.stat().st_size / 1024
                    })
    
    return problemas

def verificar_estrutura(projeto_path):
    """Verifica se tem arquivos essenciais"""
    
    projeto = Path(projeto_path)
    essenciais = {
        'README.md': False,
        'scripts': False,
    }
    
    for arquivo in essenciais:
        if (projeto / arquivo).exists():
            essenciais[arquivo] = True
    
    return essenciais

def main():
    print("=" * 80)
    print("🔒 VALIDADOR PRÉ-GITHUB - Verifica segurança LGPD")
    print("=" * 80)
    print()
    
    root = Path('.')
    total_problemas = 0
    
    for projeto in PROJETOS:
        projeto_path = root / projeto
        
        if not projeto_path.exists():
            print(f"⏭️  {projeto} - NÃO ENCONTRADO")
            print()
            continue
        
        print(f"🔍 Analisando: {projeto}")
        print("-" * 80)
        
        # Verifica estrutura
        essenciais = verificar_estrutura(projeto_path)
        print(f"  📋 Estrutura: ", end="")
        if all(essenciais.values()):
            print("✅ OK")
        else:
            for arquivo, presente in essenciais.items():
                status = "✅" if presente else "❌"
                print(f"{status} {arquivo}")
        
        # Verifica segurança
        problemas = verificar_seguranca(projeto_path)
        
        if problemas:
            print(f"  🚨 PROBLEMAS DETECTADOS ({len(problemas)}): ")
            for prob in problemas:
                if prob['tipo'] == 'pasta_perigosa':
                    print(f"     ⚠️  Pasta: {prob['path']} "
                          f"({prob['arquivos']} arquivos)")
                elif prob['tipo'] == 'arquivo_sensivel':
                    print(f"     ⚠️  Arquivo: {prob['path']} "
                          f"({prob['tamanho_kb']:.1f} KB)")
            total_problemas += len(problemas)
        else:
            print(f"  ✅ Segurança: Nenhum arquivo sensível detectado")
        
        # Verifica .gitignore
        if (projeto_path / '.gitignore').exists():
            print(f"  ✅ .gitignore: Presente")
        else:
            print(f"  ❌ .gitignore: AUSENTE")
        
        # Verifica requirements.txt
        if (projeto_path / 'requirements.txt').exists():
            print(f"  ✅ requirements.txt: Presente")
        else:
            print(f"  ⚠️  requirements.txt: Ausente (crie com gerar_requirements.py)")
        
        print()
    
    # Resumo final
    print("=" * 80)
    if total_problemas == 0:
        print("✅ SEGURO PARA FAZER PUSH!")
        print()
        print("Próximos passos:")
        print("  1. git init (se ainda não fez)")
        print("  2. git add .")
        print("  3. git commit -m 'Add data science projects - LGPD compliant'")
        print("  4. git branch -M main")
        print("  5. git push -u origin main")
    else:
        print(f"🚨 {total_problemas} PROBLEMA(S) ENCONTRADO(S)!")
        print()
        print("Execute primeiro:")
        print("  python preparar_github.py")
        print()
        print("E depois rode este validador novamente!")
    
    print("=" * 80)

if __name__ == '__main__':
    main()
