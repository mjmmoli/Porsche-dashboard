# 🏎 Porsche Sales Intelligence Dashboard

> Projeto prático desenvolvido no curso **AI Reports** — Módulo 02  
> Pipeline completo de dados: da base bruta ao dashboard interativo publicado na web

---

## Visão Geral

Este projeto demonstra como construir um pipeline de dados do zero utilizando **Python**, **Claude AI** e boas práticas de engenharia de dados — sem depender de ferramentas de BI pagas. O resultado é um dashboard 100% estático, filtrado em tempo real, publicado via GitHub Pages.

```
Planilha bruta (.xlsx)
        ↓
   Schema de regras
        ↓
  Script de higienização (Python)
        ↓
  Planilha tratada (.xlsx)
        ↓
  Dashboard interativo (HTML)
        ↓
  GitHub Pages (web pública)
```

---

## Estrutura do Projeto

```
porsche-dashboard/
├── porsche_database.xlsx          # Base de dados bruta
├── porsche_database_sanitized.xlsx # Base tratada (output)
├── sanitize_porsche.py            # Agente de higienização
├── schema.md                      # Regras de higienização
├── index.html                     # Dashboard dark (elegante)
├── index_linkedin.html            # Dashboard light (LinkedIn-ready)
└── README.md                      # Este arquivo
```

---

## A Base de Dados

### Colunas de entrada

| Coluna | Tipo | Exemplo de valor bruto |
|---|---|---|
| `sale_id` | ID único | `1`, `86` |
| `sale_date` | Data bruta | `2024-13-05`, `15 de julho de 2024` |
| `customer_name` | Nome do cliente | `John Smith`, `MARY johnson` |
| `porsche_model` | Modelo do veículo | `911 carrera`, `macan electric` |
| `model_year` | Ano do modelo | `2023`, `20-24`, `dois mil e vinte e dois` |
| `sale_price` | Preço em formatos variados | `$985.000,00`, `188 mil USD` |
| `vehicle_mileage` | Quilometragem | `12.500 milhas`, `15.200 km`, `novo` |
| `payment_method` | Método de pagamento | `credit card`, `wire`, `cripto` |
| `city` | Cidade | `new york`, `SAN JOSE` |
| `state` | Estado dos EUA | `ny`, `California`, `ca` |
| `salesperson` | Vendedor | `AMANDA scott`, `kevin` |
| `delivery_status` | Status | `Entregue!!!`, `DELIVERD`, `in-transit` |

### Problemas encontrados na base bruta

- Datas em formatos inconsistentes (ISO, americano, europeu, por extenso)
- Datas de calendário impossíveis (`2024-13-05`, `31 de abril`)
- Modelos escritos em caixa baixa, com abreviações ou erros de digitação
- Anos do modelo em texto por extenso (`dois mil e vinte e dois`)
- Preços com símbolos, sufixos (`k`, `mil`, `USD`) e separadores mistos
- Quilometragem em km precisando conversão para milhas
- Estados dos EUA como nome completo, abreviação ou caixa incorreta
- Status de entrega com pontuação, maiúsculas e erros ortográficos (`DELIVERD`)

---

## Schema de Higienização

O arquivo `schema.md` define as regras de normalização para cada coluna.

### Datas → `SaleDateSanitized`

Normalizar para o formato ISO `AAAA-MM-DD`. Aceitar:

- `AAAA-MM-DD` / `AAAA/MM/DD` / `AAAA.MM.DD`
- `MM/DD/AAAA` / `MM/DD/AA` / `MM-DD-AA`
- `Mês DD, AAAA` / `Mês DD de AAAA`

Datas impossíveis como `2024-13-05` ou `31 de abril` → `INVALIDO`

### Modelos → `PorscheModelSanitized`

Normalizar para o nome canônico da Porsche. Exemplos:

| Entrada | Saída |
|---|---|
| `911 carrera` | `911 Carrera` |
| `macan electric` | `Macan Electric` |
| `CAYENNE turbo gt` | `Cayenne Turbo GT` |
| `718 cayman gt4 rs` | `718 Cayman GT4 RS` |

Modelos desconhecidos recebem title case, sem ser descartados.

### Ano do modelo → `ModelYearSanitized`

| Entrada | Saída |
|---|---|
| `2024` | `2024` |
| `20-24` | `2024` |
| `dois mil e vinte e dois` | `2022` |
| `1985` | `INVALIDO` (fora do range 1990–2035) |

### Preço de venda → `SalesPriceSanitized`

| Entrada | Saída |
|---|---|
| `$ 985.000,00` | `985000.00` |
| `188 mil USD` | `188000.00` |
| `$645k` | `645000.00` |
| `oitenta e dois mil` | `82000.00` |

### Quilometragem → `VehicleMileageSanitized`

| Entrada | Saída |
|---|---|
| `12.500 milhas` | `12500` |
| `15.200 km` | `9444` (convertido: × 0,621371) |
| `novo` / `zero miles` | `0` |
| `doze mil milhas` | `12000` |

### Método de pagamento → `PayMethodSanitized`

Valores aceitos: `Cartão de Crédito`, `Cartão de Débito`, `Transferência Bancária`, `Transferência Eletrônica`, `Financiamento`, `Leasing`, `Dinheiro`, `Pagamento ACH`, `Pagamento em Criptomoedas`

### Estado → `StateSanitized`

| Entrada | Saída |
|---|---|
| `california` | `CA` |
| `New York` | `NY` |
| `ny` | `NY` |
| `Estado desconhecido` | `INVALIDO` |

### Status de entrega → `DeliveryStatusSanitized`

Valores aceitos: `Entregue`, `Pendente`, `Em Trânsito`, `Cancelado`, `Aguardando Entrega`, `Aguardando Recolhimento`, `Aprovação Pendente`, `Revisão Pendente`, `Enviado`

Corrige automaticamente: `DELIVERD` → `Entregue`, `Entregue!!!` → `Entregue`

---

## O Script de Higienização

### Arquivo: `sanitize_porsche.py`

O script lê a planilha bruta, aplica todas as regras do schema e gera uma nova planilha com as colunas sanitizadas inseridas ao lado das originais.

**Uso:**

```bash
pip install openpyxl
python sanitize_porsche.py porsche_database.xlsx porsche_database_sanitized.xlsx
```

**Saída no terminal:**

```
OK -> porsche_database_sanitized.xlsx
Contagem de INVALIDOS por coluna sanitizada:
  SaleDateSanitized: 24
```

### Como funciona o pipeline

```python
PIPELINE_COLUNAS = [
    ("sale_date",       "SaleDateSanitized",       sanitizar_data),
    ("porsche_model",   "PorscheModelSanitized",   sanitizar_modelo),
    ("model_year",      "ModelYearSanitized",      sanitizar_ano),
    ("sale_price",      "SalesPriceSanitized",     sanitizar_preco),
    ("vehicle_mileage", "VehicleMileageSanitized", sanitizar_quilometragem),
    ("payment_method",  "PayMethodSanitized",      sanitizar_pagamento),
    ("city",            "CitySanitized",           sanitizar_cidade),
    ("state",           "StateSanitized",          sanitizar_estado),
    ("delivery_status", "DeliveryStatusSanitized", sanitizar_entrega),
]
```

Cada função de sanitização é independente e retorna `INVALIDO` quando o valor não pode ser normalizado com segurança — nunca deixa o campo em branco.

### Resultado da higienização

| Métrica | Valor |
|---|---|
| Total de registros | 100 |
| Colunas sanitizadas | 9 |
| Datas inválidas | 24 (calendário impossível) |
| Demais colunas | 0 inválidos |

---

## O Dashboard

### Funcionalidades

- **4 KPIs** em tempo real: total de vendas, receita, ticket médio e modelo líder
- **Filtros combinados**: modelo, ano do modelo, cidade e método de pagamento
- **Ranking de modelos** com barras proporcionais
- **Vendas por ano** com colunas comparativas
- **Top modelos por cidade** com tabela ordenada por volume
- **Distribuição de pagamento** com gráfico donut
- **Preferência por gênero** inferida a partir do primeiro nome
- **Insights por cidade** com modelo favorito e ticket médio

### Duas versões disponíveis

| Versão | Arquivo | Estilo |
|---|---|---|
| Dark / elegante | `index.html` | Fundo escuro, paleta aço e dourado |
| Light / LinkedIn | `index_linkedin.html` | Fundo branco, paleta vibrante |

---

## Tecnologias utilizadas

| Tecnologia | Função |
|---|---|
| Python 3 | Script de higienização |
| openpyxl | Leitura e escrita de `.xlsx` |
| Chart.js | Gráfico donut interativo |
| HTML + CSS + JS | Dashboard estático |
| GitHub Pages | Publicação gratuita na web |
| Claude AI | Geração e revisão do código |

---

## Como publicar no GitHub Pages

1. Crie um repositório público no [github.com](https://github.com)
2. Faça upload do arquivo `index.html`
3. Vá em **Settings → Pages → Deploy from a branch → main → / (root)**
4. Aguarde 1–2 minutos e acesse:

```
https://seu-usuario.github.io/porsche-dashboard
```

---

## Aprendizados do projeto

- Como projetar um schema de higienização robusto para dados reais e sujos
- Como construir um agente Python que sanitiza dados automaticamente com regras claras
- Como transformar uma planilha tratada em um dashboard 100% estático e interativo
- Como usar Claude AI como co-piloto em cada etapa do pipeline

---

## Curso

Projeto desenvolvido no **Módulo 02 — Criando Agentes de Tratamento de Dados e Dashboards com Excel e Claude Code**, parte do curso **AI Reports**.

> Ferramentas: Python · Claude Code · Excel · GitHub Pages

---

*Construído com Python + Claude AI · Curso AI Reports*
