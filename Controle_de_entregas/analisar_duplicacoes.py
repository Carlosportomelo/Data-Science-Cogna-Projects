"""
Análise de Duplicações e Proposta de Reorganização
Identifica bases duplicadas e sugere estrutura otimizada
"""

import os
import hashlib
from pathlib import Path
from collections import defaultdict
import pandas as pd
from datetime import datetime

ROOT_DIR = r"C:\Users\a483650\Projetos"

def calcular_hash(filepath, tamanho_bloco=65536):
    """Calcula hash MD5 de arquivo para identificar duplicações exatas"""
    try:
        hasher = hashlib.md5()
        with open(filepath, 'rb') as f:
            buf = f.read(tamanho_bloco)
            while len(buf) > 0:
                hasher.update(buf)
                buf = f.read(tamanho_bloco)
        return hasher.hexdigest()
    except:
        return None

def analisar_duplicacoes():
    """Analisa arquivos duplicados por hash"""
    print("🔍 Analisando duplicações de arquivos...\n")
    
    # Dicionário: hash -> [lista de arquivos]
    arquivos_por_hash = defaultdict(list)
    arquivos_por_nome = defaultdict(list)
    
    extensoes_interesse = {'.csv', '.xlsx', '.xls'}
    total_arquivos = 0
    tamanho_total = 0
    
    # Scan de todos os arquivos
    for root, dirs, files in os.walk(ROOT_DIR):
        # Ignorar pastas .venv, __pycache__, .git
        dirs[:] = [d for d in dirs if d not in {'.venv', '__pycache__', '.git', 'venv', 'node_modules'}]
        
        for file in files:
            ext = Path(file).suffix.lower()
            if ext in extensoes_interesse:
                filepath = os.path.join(root, file)
                try:
                    tamanho = os.path.getsize(filepath)
                    tamanho_total += tamanho
                    total_arquivos += 1
                    
                    # Hash para duplicações exatas
                    if tamanho < 100 * 1024 * 1024:  # Só calcular hash de arquivos < 100MB
                        file_hash = calcular_hash(filepath)
                        if file_hash:
                            arquivos_por_hash[file_hash].append({
                                'path': filepath,
                                'size': tamanho,
                                'name': file,
                                'modified': datetime.fromtimestamp(os.path.getmtime(filepath))
                            })
                    
                    # Agrupamento por nome
                    nome_base = file.lower()
                    arquivos_por_nome[nome_base].append({
                        'path': filepath,
                        'size': tamanho,
                        'modified': datetime.fromtimestamp(os.path.getmtime(filepath))
                    })
                    
                    if total_arquivos % 100 == 0:
                        print(f"   Processados: {total_arquivos} arquivos...")
                except Exception as e:
                    pass
    
    print(f"\n✅ Análise concluída!")
    print(f"   Total de arquivos: {total_arquivos}")
    print(f"   Tamanho total: {tamanho_total / (1024**3):.2f} GB\n")
    
    return arquivos_por_hash, arquivos_por_nome, total_arquivos, tamanho_total

def gerar_relatorio_duplicacoes(arquivos_por_hash, arquivos_por_nome):
    """Gera relatório de duplicações"""
    
    # Duplicações exatas (mesmo conteúdo)
    duplicados_exatos = {h: files for h, files in arquivos_por_hash.items() if len(files) > 1}
    
    # Duplicações por nome (podem ser diferentes versões)
    duplicados_nome = {n: files for n, files in arquivos_por_nome.items() if len(files) > 1}
    
    print("="*80)
    print("📊 RELATÓRIO DE DUPLICAÇÕES")
    print("="*80)
    
    # Duplicações exatas
    print(f"\n🎯 DUPLICAÇÕES EXATAS (mesmo conteúdo):")
    print(f"   Total de grupos: {len(duplicados_exatos)}")
    
    espaco_desperdicado = 0
    detalhes_duplicados = []
    
    for hash_val, files in sorted(duplicados_exatos.items(), key=lambda x: x[1][0]['size'], reverse=True)[:20]:
        tamanho = files[0]['size']
        qtd_copias = len(files)
        desperdicio = tamanho * (qtd_copias - 1)
        espaco_desperdicado += desperdicio
        
        detalhes_duplicados.append({
            'Nome': files[0]['name'],
            'Cópias': qtd_copias,
            'Tamanho_MB': round(tamanho / (1024**2), 2),
            'Desperdício_MB': round(desperdicio / (1024**2), 2),
            'Locais': ' | '.join([f['path'].replace(ROOT_DIR, '...') for f in files[:3]])
        })
    
    print(f"   Espaço desperdiçado: {espaco_desperdicado / (1024**2):.2f} MB\n")
    
    # Top 20 duplicações
    print("\n🔝 TOP 20 ARQUIVOS DUPLICADOS:")
    for i, d in enumerate(detalhes_duplicados[:20], 1):
        print(f"\n{i}. {d['Nome']}")
        print(f"   Cópias: {d['Cópias']} | Tamanho: {d['Tamanho_MB']} MB | Desperdício: {d['Desperdício_MB']} MB")
        print(f"   Locais: {d['Locais'][:150]}...")
    
    # Análise de bases críticas
    print("\n\n" + "="*80)
    print("🎯 ANÁLISE DE BASES CRÍTICAS")
    print("="*80)
    
    bases_criticas = [
        'hubspot_leads.csv',
        'hubspot_negocios_perdidos.csv',
        'matriculas_finais',
        'meta_dataset.csv',
        'google_dataset.csv'
    ]
    
    for base in bases_criticas:
        ocorrencias = [files for nome, files in arquivos_por_nome.items() if base.lower() in nome.lower()]
        if ocorrencias:
            total_ocorrencias = sum(len(f) for f in ocorrencias)
            print(f"\n📌 {base.upper()}")
            print(f"   Ocorrências encontradas: {total_ocorrencias}")
            
            for grupo in ocorrencias[:1]:  # Mostrar um grupo
                print(f"   Variações:")
                for arq in sorted(grupo, key=lambda x: x['modified'], reverse=True)[:5]:
                    caminho_rel = arq['path'].replace(ROOT_DIR, '...')
                    tamanho_mb = arq['size'] / (1024**2)
                    print(f"      - {caminho_rel}")
                    print(f"        {tamanho_mb:.2f} MB | Modificado: {arq['modified'].strftime('%d/%m/%Y %H:%M')}")
    
    return detalhes_duplicados, espaco_desperdicado

def propor_reorganizacao():
    """Propõe estrutura de organização otimizada"""
    
    print("\n\n" + "="*80)
    print("🏗️  PROPOSTA DE REORGANIZAÇÃO")
    print("="*80)
    
    proposta = """
📁 ESTRUTURA PROPOSTA:

C:\\Users\\a483650\\Projetos\\
│
├── 📂 _DADOS_CENTRALIZADOS/              ⭐ NOVO - Repositório central de dados
│   ├── hubspot/
│   │   ├── hubspot_leads_ATUAL.csv       (Versão mais recente)
│   │   ├── hubspot_negocios_perdidos_ATUAL.csv
│   │   └── historico/                     (Backups por data)
│   ├── matriculas/
│   │   ├── matriculas_finais_ATUAL.csv
│   │   └── historico/
│   ├── marketing/
│   │   ├── meta_ads_ATUAL.csv
│   │   ├── google_ads_ATUAL.csv
│   │   └── historico/
│   └── README.md                          (Documentação das bases)
│
├── 📂 _SCRIPTS_COMPARTILHADOS/           ⭐ NOVO - Scripts reutilizáveis
│   ├── utils/
│   │   ├── conectores_hubspot.py
│   │   ├── limpeza_dados.py
│   │   └── validadores.py
│   └── README.md
│
├── 📂 01_Helio_ML_Producao/              ✅ MANTÉM (Produção)
│   ├── Data/ → SYMLINK para _DADOS_CENTRALIZADOS
│   ├── Scripts/
│   ├── Outputs/
│   └── models/
│
├── 📂 02_Pipeline_Midia_Paga/            ✅ MANTÉM (Produção)
│   ├── data/ → SYMLINK para _DADOS_CENTRALIZADOS
│   ├── scripts/
│   └── outputs/
│
├── 📂 03_Analises_Operacionais/          ⭐ CONSOLIDAR
│   ├── eficiencia_canal/
│   ├── comparativo_unidades/
│   ├── curva_alunos/
│   └── valor_the_news/
│
├── 📂 04_Auditorias_Qualidade/           ⭐ CONSOLIDAR
│   ├── auditoria_leads_sumidos/
│   ├── correcao_meta_callcenter/
│   └── integridade_bases/
│
├── 📂 05_Pesquisas_Educacionais/         ⭐ CONSOLIDAR
│   └── correlacao_cefr/
│
├── 📂 _ARQUIVO/                          ⭐ NOVO - Projetos inativos
│   ├── ambiente_teste_2025-12-15/
│   ├── projeto_helio_teste/
│   ├── analise_performance_rvo/         (vazio)
│   └── analise_performance_organica/    (vazio)
│
└── 📂 Controle_de_entregas/              ✅ MANTÉM
    ├── INVENTARIO_PROJETOS_DATA_ANALYTICS.xlsx
    └── scripts/

═══════════════════════════════════════════════════════════════════════════

🎯 BENEFÍCIOS DA REORGANIZAÇÃO:

1. ✅ Eliminação de 90% das duplicações de bases
2. ✅ Redução de ~1.2 GB de espaço (atual 1.42 GB → ~200 MB)
3. ✅ Fonte única de verdade para cada base
4. ✅ Histórico organizado por data
5. ✅ Scripts reutilizáveis entre projetos
6. ✅ Nomenclatura padronizada (01_, 02_, etc)
7. ✅ Separação clara: Produção vs Análises vs Arquivo

═══════════════════════════════════════════════════════════════════════════

⚙️  IMPLEMENTAÇÃO SUGERIDA:

FASE 1: Preparação (1h)
   1. Criar estrutura _DADOS_CENTRALIZADOS/
   2. Identificar versão mais recente de cada base
   3. Copiar bases atuais para repositório central
   4. Mover backups para /historico/

FASE 2: Consolidação (2h)
   5. Reorganizar projetos em categorias (01_, 02_, etc)
   6. Criar symlinks/atalhos para _DADOS_CENTRALIZADOS/
   7. Mover projetos inativos para _ARQUIVO/
   8. Extrair scripts comuns para _SCRIPTS_COMPARTILHADOS/

FASE 3: Validação (1h)
   9. Testar projetos em produção (Helio, Mídia Paga)
   10. Validar caminhos dos scripts
   11. Atualizar documentação
   12. Backup completo antes de deletar duplicatas

FASE 4: Limpeza (30min)
   13. Deletar arquivos duplicados
   14. Limpar pastas vazias
   15. Atualizar inventário

═══════════════════════════════════════════════════════════════════════════
"""
    
    print(proposta)
    
    # Gerar script de reorganização
    script_reorganizacao = """
# SCRIPT DE REORGANIZAÇÃO AUTOMÁTICA
# Execute por etapas com cuidado!

# FASE 1: Criar estrutura central
New-Item -Path "C:\\Users\\a483650\\Projetos\\_DADOS_CENTRALIZADOS" -ItemType Directory
New-Item -Path "C:\\Users\\a483650\\Projetos\\_DADOS_CENTRALIZADOS\\hubspot" -ItemType Directory
New-Item -Path "C:\\Users\\a483650\\Projetos\\_DADOS_CENTRALIZADOS\\hubspot\\historico" -ItemType Directory
New-Item -Path "C:\\Users\\a483650\\Projetos\\_DADOS_CENTRALIZADOS\\matriculas" -ItemType Directory
New-Item -Path "C:\\Users\\a483650\\Projetos\\_DADOS_CENTRALIZADOS\\matriculas\\historico" -ItemType Directory
New-Item -Path "C:\\Users\\a483650\\Projetos\\_DADOS_CENTRALIZADOS\\marketing" -ItemType Directory
New-Item -Path "C:\\Users\\a483650\\Projetos\\_DADOS_CENTRALIZADOS\\marketing\\historico" -ItemType Directory
New-Item -Path "C:\\Users\\a483650\\Projetos\\_SCRIPTS_COMPARTILHADOS" -ItemType Directory
New-Item -Path "C:\\Users\\a483650\\Projetos\\_ARQUIVO" -ItemType Directory

# FASE 2: Copiar bases mais recentes (MANUALMENTE - revisar antes!)
# Identificar manualmente as versões mais recentes de:
# - hubspot_leads.csv
# - hubspot_negocios_perdidos.csv
# - matriculas_finais.csv
# - meta_dataset.csv
# - google_dataset.csv

# FASE 3: Renomear projetos principais
# Rename-Item -Path "C:\\Users\\a483650\\Projetos\\projeto_helio" -NewName "01_Helio_ML_Producao"
# Rename-Item -Path "C:\\Users\\a483650\\Projetos\\analise_performance_midiapaga" -NewName "02_Pipeline_Midia_Paga"

# FASE 4: Mover para arquivo
# Move-Item -Path "C:\\Users\\a483650\\Projetos\\AMBIENTE_TESTE_ISOLADO_2025-12-15" -Destination "C:\\Users\\a483650\\Projetos\\_ARQUIVO\\"
# Move-Item -Path "C:\\Users\\a483650\\Projetos\\projeto_helio_teste" -Destination "C:\\Users\\a483650\\Projetos\\_ARQUIVO\\"
"""
    
    with open(r"C:\Users\a483650\Projetos\Controle_de_entregas\script_reorganizacao.ps1", 'w', encoding='utf-8') as f:
        f.write(script_reorganizacao)
    
    print("\n💾 Script PowerShell salvo: script_reorganizacao.ps1")

# Executar análise
print("="*80)
print("🚀 INICIANDO ANÁLISE DE DUPLICAÇÕES")
print("="*80)

arquivos_por_hash, arquivos_por_nome, total_arquivos, tamanho_total = analisar_duplicacoes()
detalhes, espaco_desperdicado = gerar_relatorio_duplicacoes(arquivos_por_hash, arquivos_por_nome)

# Salvar relatório em Excel
df_duplicados = pd.DataFrame(detalhes)
output_path = r"C:\Users\a483650\Projetos\Controle_de_entregas\RELATORIO_DUPLICACOES.xlsx"
df_duplicados.to_excel(output_path, index=False)
print(f"\n💾 Relatório Excel salvo: {output_path}")

propor_reorganizacao()

print("\n" + "="*80)
print("✅ ANÁLISE CONCLUÍDA!")
print("="*80)
print(f"\n📊 Estatísticas Finais:")
print(f"   Total de arquivos analisados: {total_arquivos}")
print(f"   Tamanho total: {tamanho_total / (1024**3):.2f} GB")
print(f"   Espaço desperdiçado com duplicatas: {espaco_desperdicado / (1024**2):.2f} MB")
print(f"   Potencial de economia: {(espaco_desperdicado / tamanho_total * 100):.1f}%")
print(f"\n📁 Próximos passos:")
print(f"   1. Revisar RELATORIO_DUPLICACOES.xlsx")
print(f"   2. Analisar proposta de reorganização")
print(f"   3. Fazer backup completo antes de executar")
print(f"   4. Executar script_reorganizacao.ps1 por fases")
