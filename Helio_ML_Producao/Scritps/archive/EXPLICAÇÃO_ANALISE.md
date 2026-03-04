# 📚 Documentação Técnica: Scripts de Análise de Recuperação

**Projeto:** Análise de Recuperação Alta 26.1 (Red Balloon)
**Data:** 11/12/2025
**Contexto:** Preparação de dados (Data Prep) para modelos de Classificação e Regressão.

---

## 🛠️ Resumo das Ferramentas

Atualmente, dispomos de 4 scripts principais na pasta `Scripts`. Cada um tem uma função específica no pipeline de engenharia de dados.

| Script | Função Principal | Quando usar? |
| :--- | :--- | :--- |
| **1. `Data_check.py`** | **Carregamento e Limpeza** | Sempre que baixar um CSV novo e quiser garantir que ele abre sem erro de encoding ou colunas erradas. |
| **2. `data_audit.py`** | **Auditoria de Volume** | Para comparar se o número total de linhas do CSV bate com o número total do Dashboard (PowerBI). |
| **3. `check_fontes_detalhadas.py`** | **Raio-X de Strings** | Para descobrir como as plataformas (Google, Facebook) estão escritas "sujas" dentro das colunas de texto. |
| **4. `debug_dash.py`** | **Reconciliação (Tira-Teima)** | Script avançado para quebrar os dados por Mês/Safra e descobrir qual filtro o Dashboard original está usando. |

---

## 🔍 Detalhamento Técnico

### 1. `Data_check.py` (O Porteiro)
Este script resolve o problema clássico de **"KeyError"** e **"UnicodeDecodeError"**.
* **O que ele faz:**
    * Tenta ler o arquivo em `UTF-8`. Se falhar, tenta `Latin-1` (padrão Excel Brasil).
    * Remove espaços vazios dos nomes das colunas (ex: `"Data "` vira `"Data"`).
    * Renomeia as colunas técnicas do HubSpot para português legível (ex: `Data de criaÃ§Ã£o` -> `Data de Criação`).
    * Verifica se a coluna de Data é válida.
* **Saída:** Um resumo no terminal dizendo "Dados carregados com sucesso" e mostrando o período (Data Min/Max).

### 2. `data_audit.py` (O Auditor)
Este script serve para garantir a **Sanidade dos Dados**.
* **O que ele faz:**
    * Agrupa os dados por **ANO** (para vermos a evolução histórica).
    * Tenta replicar as categorias de negócio (Digital Pago vs Orgânico).
    * Calcula taxas macro de conversão.
* **O Insight que ele gerou:** Descobrimos que o CSV contém todo o histórico (102k leads), enquanto a imagem de referência (Dashboard) mostra apenas um recorte (18k leads).

### 3. `check_fontes_detalhadas.py` (A Lupa)
Este script é essencial para a **Engenharia de Features** do modelo de IA.
* **O que ele faz:**
    * Olha dentro da coluna `Detalhamento da fonte original do tráfego 1`.
    * Lista todas as variações de nomes de campanha (ex: `red balloon | google search | sp`).
* **Por que é importante:** O modelo precisa saber separar **Google (Busca)** de **Facebook (Social)**. Como os dados vêm "sujos" com nomes de campanhas, este script nos dá a lista para criar o "Dicionário de Tradução".

### 4. `debug_dash.py` (O Detetive)
*Status: Em uso atual*
* **O que ele faz:**
    * Cria uma visão pivotada (Tabela Dinâmica) de **Mês x Categoria**.
* **Objetivo:** Encontrar em qual mês/ano os números somam **18.781** (o número mágico do Dashboard). Isso nos permitirá filtrar a base corretamente para treinar a IA apenas no cenário atual, sem usar dados obsoletos de 2021.

---

## 🚀 Como Executar
No PowerShell, utilize o interpretador Python apontando para o script desejado:

```powershell
# Exemplo para rodar a auditoria
& python.exe c:/caminho/do/projeto/Scripts/data_audit.py

⚠️ Pontos de Atenção (Data Quality)

    Case Sensitive: O Python diferencia Google de google. Sempre usamos .lower() nos scripts para evitar erros.

    Encoding: O HubSpot geralmente exporta em UTF-8, mas se salvar no Excel, vira ANSI/Latin-1. Os scripts já tratam isso automaticamente.