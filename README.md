# 🏎️ Dashboard de Vendas e Performance Porsche (HTML & Excel)

<div align="center">

![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![CSS3](https://img.shields.io/badge/CSS3-1572B6?style=for-the-badge&logo=css3&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![Chart.js](https://img.shields.io/badge/Chart.js-FF6384?style=for-the-badge&logo=chartdotjs&logoColor=white)
![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Porsche](https://img.shields.io/badge/Porsche-Motorsport-D5001C?style=for-the-badge&logo=porsche&logoColor=white)
![Status](https://img.shields.io/badge/Status-Concluido-00C853?style=for-the-badge)
![DIO](https://img.shields.io/badge/DIO-Santander_Coders-E94D5F?style=for-the-badge)

**Dashboard Executivo Interativo de Vendas Porsche em HTML/JS com Pipeline Automatizado de Sanitizacao em Python e Excel**

[📌 Sobre](#-sobre-o-projeto) • [🎯 Objetivos](#-objetivos-do-projeto) • [🧹 Higienizacao](#-regras-de-higienizacao-de-dados) • [📊 Painel HTML](#-recursos-do-dashboard-html) • [🚀 Como Executar](#-como-executar) • [📂 Estrutura](#-estrutura-de-arquivos)

</div>

---

## 📌 Sobre o Projeto

Este projeto foi desenvolvido na trilha **Santander - Excel com IA e Claude (Aceleracao AI Reports com Excel, GPT Agents e Claude Code)** pela **DIO (Digital Innovation One)**.

Ele implementa uma solucao integrada de **Engenharia e Visualizacao de Dados**:
1. **Pipeline de Sanitizacao Automatizado em Python (`Script Planilha Porsche2.py`)**: Le a base bruta `porsche_database.xlsx` com dados despadronizados de vendas, aplica regras estritas de higienizacao e valida calendarizacao, moedas, milhagens, cidades e estados norte-americanos.
2. **Workbook Saneado (`porsche_database_sanitized.xlsx`)**: Estrutura analitica com cada coluna tratada inserida imediatamente apos a coluna fonte, rotulada com sufixo `Sanitized` e celulas invalidas identificadas como `INVALID`.
3. **Dashboard Web Executivo (`dashboard-porsche.html`)**: Interface moderna em tema escuro com identidade visual da Porsche (Preto, Branco, Vermelho Motorsport e Dourado), gerando KPIs em tempo real, graficos interativos via Chart.js, filtros multidimensionais e exportacao instantanea para XLSX via SheetJS.

---

## 🎯 Objetivos do Projeto

<div align="center">

| # | 🎯 Pilar | 📋 Entrega Realizada |
|---|---|---|
| 1 | 🧹 **Saneamento de Dados** | Transformacao de precos textuais/mistos, datas invalidas, milhas e siglas de estados |
| 2 | 📊 **KPIs Executivos (Big Numbers)** | Faturamento total, ticket medio, unidades entregues e modelo lider de vendas |
| 3 | 📈 **Visualizacao Interativa** | Graficos de barras horizontais/verticais e distribuicao de status em doughnut |
| 4 | 🎛️ **Filtros Dinamicos** | Filtros combinados por Modelo, Ano de Fabricacao, Status de Entrega, Pagamento e Busca Livre |
| 5 | 📥 **Exportacao de Dados** | Geracao sob demanda de relatorios Excel (`.xlsx`) com os dados filtrados em tela |

</div>

---

## 🧹 Regras de Higienizacao de Dados (Schema)

Conforme o schema definido em `Prompt Schema Porsche.md`, cada registro foi normalizado:

| Coluna Original | Coluna Sanitizada | Regras Aplicadas |
|---|---|---|
| `sale_date` | `SaleDateSanitized` | Formato ISO `YYYY-MM-DD`. Datas impossiveis (ex.: 30 de fevereiro, 31 de abril) convertidas em `INVALID` |
| `porsche_model` | `PorscheModelSanitized` | Title Case canonico preservando siglas tecnicas (`911 Turbo S`, `718 Cayman GT4 RS`, `E-Hybrid`) |
| `model_year` | `ModelYearSanitized` | 4 digitos inteiros (`1990` a `2035`). Converte extenso (`twenty twenty four` -> `2024`) |
| `sale_price` | `SalesPriceSanitized` | Decimal USD com duas casas (`.2f`). Remove `$`, `USD`, `dollars` e interpreta multiplicadores `k` |
| `vehicle_mileage` | `VehicleMileageSanitized` | Inteiro em milhas. Converte `KM` para milhas ($1\text{ km} \approx 0.621371\text{ mi}$) e zera `new`/`zero miles` |
| `payment_method` | `PayMethodSanitized` | Lista controlada: `Credit Card`, `Wire Transfer`, `Financing`, `Cash`, `Lease`, etc. |
| `city` | `CitySanitized` | Title Case preservando pontuacoes como `St. Louis` |
| `state` | `StateSanitized` | Sigla oficial USPS com 2 letras (`CA`, `NY`, `TX`). Nomes por extenso convertidos |
| `delivery_status` | `DeliveryStatusSanitized` | Lista controlada: `Delivered`, `In Transit`, `Pending`, `Cancelled`, `Shipped`, etc. |

---

## 📊 Recursos do Dashboard HTML

```
┌─────────────────────────────────────────────────────────────────────────────┐
│ 🏎️ PORSCHE INTELLIGENCE | Executive Sales Analytics       [ 📥 EXPORTAR XLSX ] │
├─────────────────────────────────────────────────────────────────────────────┤
│ 💰 FATURAMENTO TOTAL  │ 🚗 UNIDADES VENDIDAS  │ ⚡ MILHAGEM MEDIA │ 🏆 MODELO LIDER │
│   $12,827,800.50     │      100 veiculos     │     11,422 mi     │   911 Turbo S   │
│ Ticket: $128,278.01  │    72% Entregues      │   Media frota     │ $1,880,000.00   │
├─────────────────────────────────────────────────────────────────────────────┤
│ 🎛️ FILTROS DINAMICOS:                                                       │
│ [ Modelo: Todos ]  [ Ano: Todos ]  [ Status: Todos ]  [ Busca: _________ ]   │
├─────────────────────────────────────────────┬───────────────────────────────┤
│ 📊 TOP 10 MODELOS POR FATURAMENTO           │ 🍩 DISTRIBUICAO POR STATUS    │
│ 911 Turbo S     ████████████████ ($1.88M)   │ • Entregues (Delivered): 72%  │
│ Cayenne Turbo   ████████████ ($1.42M)       │ • Em Transito (In Transit): 14%│
│ Taycan Turbo S  ██████████ ($1.15M)         │ • Pendentes (Pending): 8%     │
│ Macan GTS       ████████ ($920k)            │ • Cancelados / Outros: 6%     │
├─────────────────────────────────────────────┴───────────────────────────────┤
│ 📋 TABELA DINAMICA COM BADGES DE STATUS E BUSCA EM TEMPO REAL               │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 🗺️ Pipeline de Execucao

```mermaid
flowchart TD
    A["📥 porsche_database.xlsx\n(Dados brutos/despadronizados)"] --> B["🐍 Script Planilha Porsche2.py\n(Engine de Sanitizacao)"]
    B --> C["🧹 Aplicacao de Regras\nDatas · Precos · Modelos · UF"]
    C --> D["📗 porsche_database_sanitized.xlsx\n(Planilha Higienizada com 21 colunas)"]
    C --> E["🌐 dashboard-porsche.html\n(Dashboard Interativo Standalone)"]
    E --> F["📊 Visualizacao Executiva\nKPIs · Graficos Chart.js · Tabela"]
    E --> G["📥 Exportacao XLSX Sob Demanda\n(SheetJS Client-side)"]
```

---

## 🚀 Como Executar

### 1️⃣ Pre-requisitos

Certifique-se de possuir o Python 3.10+ e as dependencias necessarias:

```bash
pip install openpyxl
```

### 2️⃣ Executar o Script Completo

Execute o script diretamente na pasta do projeto:

```bash
python "Script Planilha Porsche2.py"
```

O comando executara:
* ✅ Higienizacao completa e salvamento de `porsche_database_sanitized.xlsx`
* ✅ Geracao automatica do dashboard atualizado `dashboard-porsche.html`

### 3️⃣ Abrir o Dashboard

Basta dar um duplo clique no arquivo `dashboard-porsche.html` ou abri-lo no seu navegador favorito (Chrome, Edge, Firefox, Safari).

---

## 📂 Estrutura de Arquivos

```
📦 Santander - Excel com IA e Claude/
 ├── 🐍 Script Planilha Porsche2.py      # Engine em Python de saneamento e gerador do HTML
 ├── 📗 porsche_database.xlsx            # Base bruta de dados original
 ├── 📗 porsche_database_sanitized.xlsx  # Planilha corrigida com colunas Sanitized
 ├── 🌐 dashboard-porsche.html           # Dashboard executivo web responsivo
 ├── 📄 Prompt Schema Porsche.md         # Documento com especificacoes e regras de saneamento
 └── 📄 README.md                        # Documentacao profissional do projeto
```

---

## 🛠️ Tecnologias Utilizadas

<div align="center">

| Tecnologia | Finalidade |
|---|---|
| ![Python](https://img.shields.io/badge/-Python-3776AB?logo=python&logoColor=white) | Logica de parsing, regex e saneamento de dados |
| ![Excel](https://img.shields.io/badge/-openpyxl-217346?logo=microsoft-excel&logoColor=white) | Leitura, estilizacao e gravacao de planilhas `.xlsx` |
| ![HTML5](https://img.shields.io/badge/-HTML5-E34F26?logo=html5&logoColor=white) | Estrutura semantica do dashboard |
| ![CSS3](https://img.shields.io/badge/-CSS3%20Flexbox%2FGrid-1572B6?logo=css3&logoColor=white) | Identidade visual escura Porsche Motorsport |
| ![JavaScript](https://img.shields.io/badge/-JavaScript%20ES6+-F7DF1E?logo=javascript&logoColor=black) | Reatividade, calculo de KPIs e filtros em tempo real |
| ![Chart.js](https://img.shields.io/badge/-Chart.js-FF6384?logo=chartdotjs&logoColor=white) | Renderizacao dos graficos de faturamento e status |
| ![SheetJS](https://img.shields.io/badge/-SheetJS%20(xlsx)-217346?logo=excel&logoColor=white) | Exportacao de dados filtrados para `.xlsx` diretamente no browser |

</div>

---

## 👤 Autor

<div align="center">

**Projeto desenvolvido com 🏎️ foco em excelencia analitica e visual.**

[![DIO](https://img.shields.io/badge/DIO-Meu_Perfil-E94D5F?style=for-the-badge)](https://www.dio.me)
[![GitHub](https://img.shields.io/badge/GitHub-100000?style=for-the-badge&logo=github&logoColor=white)](https://github.com)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://linkedin.com)

</div>

---

<div align="center">

⭐ **Se este projeto foi util para voce, considere deixar uma estrela no repositorio!** ⭐

`🏎️ Porsche` • `📊 HTML Dashboard` • `🐍 Python` • `📗 Excel Saneado` • `📈 Chart.js` • `🤖 AI Reports`

</div>