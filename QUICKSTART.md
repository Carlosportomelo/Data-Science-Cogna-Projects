# 🚀 QUICK START - 5 minutos para começar

Escolha seu projeto e siga as instruções abaixo.

---

## ⚡ Opção 1: SEM instalar nada (Just Look)

Quer só explorar o código? Use GitHub Web:

1. Clique nas pastas: `01_HELIO_ML_SCORING`, `02_PIPELINE_...`, etc
2. Leia o `README.md` de cada uma
3. Explore os scripts em `scripts/`
4. Ver documentação em `Docs/`

✅ **Feito em 2 minutos, zero setup necessário**

---

## 🔧 Opção 2: Executar localmente

### Pré-requisitos
- ✅ Python 3.8+: https://python.org
- ✅ Git instalado: https://git-scm.com
- ✅ ~2GB espaço em disco

### Setup (escolha UMA das opções)

<details>
<summary>📍 <b>Setup Opção A: Projeto de ML (Recomendado começar por aqui)</b></summary>

```bash
# 1. Clone
git clone https://github.com/seu_usuario/seu_repositorio.git
cd seu_repositorio\01_HELIO_ML_SCORING

# 2. Crie ambiente virtual
python -m venv .venv
.venv\Scripts\activate  # Windows
# or
source .venv/bin/activate  # Linux/Mac

# 3. Instale dependências
pip install -r requirements.txt

# 4. Explore o notebook (se houver)
jupyter notebook notebooks/  # opcional

# 5. Execute script de exemplo
python scripts/exemplo_basico.py
```

**Tempo:** ~5 min (incluindo download de dependências)

</details>

<details>
<summary>📍 <b>Setup Opção B: Pipeline de Dados (Mais avançado)</b></summary>

```bash
# 1. Navegue
cd seu_repositorio\02_PIPELINE_MIDIA_PAGA

# 2. Setup
python -m venv .venv
.venv\Scripts\activate

# 3. Instale
pip install -r requirements.txt

# 4. Execute (precisa de dados suas)
python scripts/validate_data.py --help
```

</details>

<details>
<summary>📍 <b>Setup Opção C: Análises (Dashboard)</b></summary>

```bash
cd seu_repositorio\03_ANALISES_OPERACIONAIS
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

# Rode análise de exemplo
python scripts/run_eficiencia_canal.py
```

</details>

---

## 📊 Próximas Etapas

### ✅ Depois de fazer setup:

1. **Leia a documentação**
   ```bash
   # Em qualquer projeto
   cat README.md
   cd Docs && cat *.md
   ```

2. **Explore o código**
   ```bash
   # Veja scripts principais
   ls scripts/
   cat scripts/main.py
   ```

3. **Rode um exemplo**
   ```bash
   python scripts/exemplo.py
   ```

4. **Adapte para seu dado**
   - Consulte schema em `Docs/DATA_SCHEMA.md`
   - Coloque seus dados em `data/`
   - Execute script no seu dado

---

## 🐛 Troubleshooting Rápido

### Erro: "Python não encontrado"
```bash
# Verifique instalação
python --version  # Deve ser 3.8+

# Se não tiver, instale: https://python.org
```

### Erro: "ModuleNotFoundError"
```bash
# Certifique-se que está no venv ativado
.venv\Scripts\activate

# Instale dependências novamente
pip install -r requirements.txt
```

### Erro: "Arquivo não encontrado"
```bash
# Verifique que está na pasta correta
cd 01_HELIO_ML_SCORING
dir  # Deve ver pasta 'scripts'
```

### Precisa de ajuda específica
- Veja `README.md` da pasta
- Consulte `Docs/` para documentação detalhada
- Você tem 4 projetos bem estruturados com tudo explicado!

---

## 🎯 Casos de Uso

**Sou Data Scientist** → Comece por `01_HELIO_ML_SCORING`
- Veja feature engineering
- Explore modelo de ML
- Adapte para seus dados

**Sou Data Engineer** → Comece por `02_PIPELINE_MIDIA_PAGA`
- ETL com APIs
- Validação de dados
- Pipelines automáticos

**Sou Líder/Executivo** → Comece por `03_ANALISES_OPERACIONAIS`
- Dashboards
- KPIs
- Análises de negócio

**Trabalho com Vendas** → Comece por `06_FUNIL_REDBALLOON`
- Análise de conversão
- Funil de vendas
- Previsões

---

## 📈 O que Esperar

✅ **Código limpo e bem estruturado**
✅ **Documentação técnica**
✅ **Exemplos funcionais**
✅ **Modelos prontos**

❌ **Dados de clientes** (removidos por LGPD)
❌ **Credenciais** (nunca em Git!)

---

## ❓ Perguntas Frequentes

**P: Posso usar isso em produção?**
R: Sim! O código é production-ready. Adapte para seus dados.

**P: Qual o nível de dificuldade?**
R: Intermediário a avançado. Expect conhecimento de Python.

**P: Como adaptar para meus dados?**
R: Cada projeto tem `DATA_SCHEMA.md` explicando formato esperado.

**P: Preciso pagar algo?**
R: Não! Código aberto, livre para usar.

---

## 🎓 Aprenda Vendo o Código

**Feature Engineering** → `01_HELIO_ML_SCORING/scripts`
**ETL & Integração** → `02_PIPELINE_MIDIA_PAGA/scripts`
**Analytics** → `03_ANALISES_OPERACIONAIS/scripts`
**Análise Estatística** → `06_FUNIL_REDBALLOON/scripts`

---

<div align="center">

**Pronto para começar?**

👉 Escolha um projeto acima e siga as instruções!

Tempo estimado: 5-10 minutos para rodar primeira vez

Questions? Check the README.md de cada projeto.

</div>

---

**Last Updated**: Março 2026 (Período: Out 2025 - Mar 2026)
**Versão**: 1.0
