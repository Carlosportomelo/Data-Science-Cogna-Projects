# GUIA: Como fazer push para GitHub

## Passo 1: preparar

```bash
cd C:\Users\a483650\Projetos
python preparar_github.py
python gerar_requirements.py
python validar_seguranca_github.py
```

## Passo 2: criar repositório no GitHub
- Repository name: `meu-portfolio-dados`
- Description: `Meus projetos de automação e ML`
- Visibility: Public ou Private

## Passo 3: push inicial

```bash
git init
git add .
git commit -m "Initial commit: Projetos de ML, ETL e análises de dados"
git remote add origin https://github.com/SEU_USUARIO/SEU_REPO_AQUI.git
git branch -M main
git push -u origin main
```

## Estrutura esperada

```text
Helio_ML_Producao/
Pipeline_Midia_Paga/
Analises_Operacionais/
Analise_Funil_RedBalloon/
README.md
```
