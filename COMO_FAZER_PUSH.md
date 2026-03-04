# 🚀 GUIA: Como Fazer Push para GitHub

## Checklist de Preparação

- [ ] Leia este guia inteiro
- [ ] Crie conta no GitHub (se não tiver): https://github.com
- [ ] Instale Git: https://git-scm.com/download/win
- [ ] Execute script de limpeza
- [ ] Execute script de validação
- [ ] Execute `git init`

---

## PASSO 1: Limpar Dados Sensíveis

```bash
cd C:\Users\a483650\Projetos
python preparar_github.py
```

**O que acontece:**
1. Script encontra arquivos sensíveis (CSV, Excel, etc)
2. Mostra o que será deletado
3. Você confirma ("s" ou "n")
4. Deleta tudo automaticamente

⚠️ **Cuidado**: Uma vez deletado, não recupera! Mas você pode usar o backup em `_BACKUP_SEGURANCA_*`

---

## PASSO 2: Gerar requirements.txt

Depois de deletar, gere os arquivos de dependências:

```bash
python gerar_requirements.py
```

Isso cria um `requirements.txt` em cada projeto com tudo que é necessário instalar.

---

## PASSO 3: Validar Segurança

Cheque se não ficou nada sensível:

```bash
python validar_seguranca_github.py
```

**Output esperado:**
```
✅ SEGURO PARA FAZER PUSH!

Próximos passos:
  1. git init
  2. git add .
  3. git commit -m '...'
  ...
```

Se aparecer problemas, execute novamente `preparar_github.py`.

---

## PASSO 4: Criar Repositório no GitHub

### 4.1 - Acesse GitHub

1. Vá para https://github.com/new
2. Faça login (ou crie conta)
3. Preencha:

```
Repository name: meu-portfolio-dados
  ↳ Ex: "Portfolio-DataScience", "Projetos-Inteligencia", etc

Description: Meus projetos de automação e ML (LGPD compliant)

Visibility: Public (para mostrar no portfólio)
           ou Private (só você vê)

Initialize repository: NÃO marque nada
```

4. Click "Create repository"

### 4.2 - GitHub vai mostrar um URL

Exemplo:
```
https://github.com/seu_usuario/meu-portfolio-dados.git
```

**Copie esta URL** - você vai usar nos próximos passos.

---

## PASSO 5: Configurar Git (PRIMEIRA VEZ APENAS)

Se é a primeira vez usando Git, configure identidade:

```bash
git config --global user.name "Seu Nome Completo"
git config --global user.email "seu_email@gmail.com"
```

---

## PASSO 6: Fazer Push para GitHub

Na pasta `C:\Users\a483650\Projetos`, execute:

```bash
# Inicia repositório local
git init

# Adiciona TODOS os arquivos preparados
git add .

# Cria primeiro commit
git commit -m "Initial commit: Projetos de ML, ETL e análises de dados (LGPD compliant)"

# Conecta ao repositório remoto (GitHub)
git remote add origin https://github.com/SEU_USUARIO/SEU_REPO_AQUI.git

# Envia para GitHub
git branch -M main
git push -u origin main
```

**❗ IMPORTANTE:** Substitua `SEU_USUARIO/SEU_REPO_AQUI` pela URL que GitHub deu

---

## PASSO 7: Verificar no GitHub

1. Acesse: https://github.com/seu_usuario/seu_repo_aqui
2. Você deve ver:
   - ✅ Pastas `01_Helio_ML_...`, `02_Pipeline_...`, etc
   - ✅ Arquivo `README.md` com instrução
   - ✅ Arquivo `.gitignore` com regras
   - ❌ NenhUM arquivo CSV/Excel (dados deletados)
   - ❌ Nenhuma pasta `data/` com dados

---

## ✂️ Como limpar dados em um projeto específico?

Se quiser ser mais seletivo, após rodar `preparar_github.py`, manualmente delete:

```bash
# Exemplo - Deletar dados do Projeto 02
cd 02_Pipeline_Midia_Paga
rmdir /s data          # Deleta pasta
rmdir /s outputs       # Deleta pasta
del *.csv              # Deleta todos CSV
```

---

## 🔄 Próximas atualizações (ao desenvolver mais)

Depois que fizer o primeiro push, para atualizar:

```bash
# 1. Faça suas mudanças no código
# 2. Stage das mudanças
git add .

# 3. Commit
git commit -m "Descrição do que mudou"

# 4. Push
git push origin main
```

---

## 🆘 Troubleshooting

### "fatal: not a git repository"
**Solução:**
```bash
# Certifique-se que está em C:\Users\a483650\Projetos
cd C:\Users\a483650\Projetos
git init
```

### "error: remote origin already exists"
**Solução:**
```bash
# Remove remote antigo
git remote remove origin

# Adiciona novo
git remote add origin https://github.com/seu_usuario/seu_repo.git
```

### "fatal: Authentication failed"
**Solução:**
1. GitHub requer token em vez de senha (desde 2021)
2. Crie um Personal Access Token em: https://github.com/settings/tokens
3. Use o token no lugar da senha

Ou use SSH (mais seguro):
```bash
git remote set-url origin git@github.com:seu_usuario/seu_repo.git
```

### Revisar o que vai fazer push
```bash
git status          # Mostra arquivos prontos
git diff            # Mostra mudanças
git log --oneline   # Mostra histórico
```

---

## 📊 Resumo do Projeto Final

```
seu-portfolio-dados/
├── 01_Helio_ML_Producao/          # ✅ Scripts de ML
├── 02_Pipeline_Midia_Paga/         # ✅ ETL & automação
├── 03_Analises_Operacionais/       # ✅ Business analytics
├── 06_Analise_Funil_RedBalloon/    # ✅ Análise de conversão
├── README.md                       # ✅ Documentação geral
├── .gitignore                      # ✅ Segurança LGPD
├── preparar_github.py              # 🔧 Ferramentas
├── gerar_requirements.py           #    de setup
├── validar_seguranca_github.py     #    auxiliares
└── COMO_FAZER_PUSH.md              # Este arquivo!
```

---

## 💡 Dicas para Portfólio

Depois de fazer o push:

1. **Personalize o GitHub Profile:**
   - Foto de perfil
   - Bio com skills
   - Link para este repositório

2. **Adicione na seção "Featured" do perfil:**
   - Para destaque maior

3. **Crie um `PORTFOLIO.md` no repositório:**
   - Descreve cada projeto
   - Links e screenshots

4. **Showcase no LinkedIn:**
   - "Check out my GitHub portfolio"
   - Link direto para o repo

---

**✅ Pronto! Você está com tudo sob controle.**

Qualquer dúvida, refaça este guia. Sucesso! 🚀
