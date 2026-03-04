
# SCRIPT DE REORGANIZAÇÃO AUTOMÁTICA
# Execute por etapas com cuidado!

# FASE 1: Criar estrutura central
New-Item -Path "C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS" -ItemType Directory
New-Item -Path "C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot" -ItemType Directory
New-Item -Path "C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\historico" -ItemType Directory
New-Item -Path "C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\matriculas" -ItemType Directory
New-Item -Path "C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\matriculas\historico" -ItemType Directory
New-Item -Path "C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\marketing" -ItemType Directory
New-Item -Path "C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\marketing\historico" -ItemType Directory
New-Item -Path "C:\Users\a483650\Projetos\_SCRIPTS_COMPARTILHADOS" -ItemType Directory
New-Item -Path "C:\Users\a483650\Projetos\_ARQUIVO" -ItemType Directory

# FASE 2: Copiar bases mais recentes (MANUALMENTE - revisar antes!)
# Identificar manualmente as versões mais recentes de:
# - hubspot_leads.csv
# - hubspot_negocios_perdidos.csv
# - matriculas_finais.csv
# - meta_dataset.csv
# - google_dataset.csv

# FASE 3: Renomear projetos principais
# Rename-Item -Path "C:\Users\a483650\Projetos\projeto_helio" -NewName "01_Helio_ML_Producao"
# Rename-Item -Path "C:\Users\a483650\Projetos\analise_performance_midiapaga" -NewName "02_Pipeline_Midia_Paga"

# FASE 4: Mover para arquivo
# Move-Item -Path "C:\Users\a483650\Projetos\AMBIENTE_TESTE_ISOLADO_2025-12-15" -Destination "C:\Users\a483650\Projetos\_ARQUIVO\"
# Move-Item -Path "C:\Users\a483650\Projetos\projeto_helio_teste" -Destination "C:\Users\a483650\Projetos\_ARQUIVO\"
