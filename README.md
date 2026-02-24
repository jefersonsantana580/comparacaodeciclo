
# Passo 1 — Comparativo Mensal por PRODUCT SERIES (REQUEST − PLAN)

Este repositório contém um script Python que lê um arquivo Excel com as abas **PLAN** e **REQUEST**,
detecta automaticamente as colunas de **meses pt‑BR** (`jan/26 … dez/26`), agrega por **`SITE` + `PRODUCT SERIES`**
e gera um **comparativo mensal (REQUEST − PLAN)**, incluindo **`TOTAL` por linha** e a linha **`TOTAL GERAL`**
(soma por coluna e total anual). O resultado é gravado na mesma planilha, na aba `Step1_Comparativo_Serie` (por padrão).

> **Observação:** O script não depende de nomes fixos de séries; ele detecta os meses pelo padrão `MMM/AA` em pt‑BR
> (`jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez`) e ordena por **mês/ano** automaticamente.

---

## 📂 Estrutura
```
.
├── step1_comparativo_serie.py   # Script principal (Passo 1)
├── requirements.txt             # Dependências (pandas, openpyxl)
└── README.md                    # Este guia
```

---

## ✅ Pré‑requisitos
- **Python 3.9+** (recomendado 3.10 ou 3.11)
- Pacotes do `requirements.txt`:
  - `pandas>=2.0.0`
  - `openpyxl>=3.1.0`

---

## 🚀 Como usar
### 1) Criar e ativar um ambiente virtual (opcional, recomendado)
```bash
python -m venv .venv
# Windows
.venv\\Scripts\\activate
# macOS / Linux
source .venv/bin/activate
```

### 2) Instalar as dependências
```bash
pip install -r requirements.txt
```

### 3) Executar o Passo 1 (gera/atualiza a aba `Step1_Comparativo_Serie`)
```bash
python step1_comparativo_serie.py --arquivo "testar _ analise de ciclo.xlsx"
```

### Parâmetros opcionais
- `--aba_saida` — nome da aba de saída (default: `Step1_Comparativo_Serie`).

Exemplo:
```bash
python step1_comparativo_serie.py \
  --arquivo "testar _ analise de ciclo.xlsx" \
  --aba_saida "Step1_Comparativo_Serie"
```

---

## 🧾 Expectativa de esquema das abas
O script espera encontrar no arquivo Excel:
- **Aba `PLAN`** e **aba `REQUEST`**;
- Colunas de **meses** no formato `MMM/AA` em pt‑BR (ex.: `jan/26`, `fev/26`, …, `dez/26`);
- Colunas **`SITE`** e **`PRODUCT SERIES`** para agregação;
- Demais colunas são ignoradas no cálculo (não precisam ser removidas).

> Se `SITE` e/ou `PRODUCT SERIES` não existirem em ambas as abas, o script aborta com erro informando as colunas ausentes.

---

## 🧮 Saída gerada
A aba de saída (**`Step1_Comparativo_Serie`**, por padrão) contém:
- **Chaves:** `SITE`, `PRODUCT SERIES`;
- **Colunas de meses:** somente os **meses** (valores em cada mês = **REQUEST − PLAN**);
- **`TOTAL` por linha:** soma dos deltas mensais por série;
- **Linha `TOTAL GERAL`:** somatório por **coluna** (todos os sites/séries) e **TOTAL** anual.

---

## ❗️Erros comuns & dicas
- **Arquivo não encontrado:** verifique o caminho passado em `--arquivo`.
- **Colunas de meses não detectadas:** confirme que os meses estão nomeados como `MMM/AA` em pt‑BR.
- **Chaves ausentes:** garanta que `SITE` **e** `PRODUCT SERIES` existem em **PLAN** e **REQUEST`**.

---

## 🧰 Desenvolvimento
- Formatação: padrão PEP8 (sugestão: usar `ruff`/`black`, opcionais).
- Testes simples: execute o script apontando para um arquivo de exemplo contendo abas `PLAN` e `REQUEST`.

---

## 📄 Licença
Defina a licença do seu repositório conforme a política da sua organização (ex.: MIT, Apache‑2.0). Se precisar, adiciono um `LICENSE`.
