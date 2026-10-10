# AMOK EQUIPMENT AS — Enterprise Financial Intelligence & Modeling Suite

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/)
[![SQLite 3 WAL](https://img.shields.io/badge/SQLite-WAL--Mode-green.svg)](https://www.sqlite.org/)
[![Power BI PBIP](https://img.shields.io/badge/Power%20BI-PBIP%20Dev%20Mode-yellow.svg)](https://powerbi.microsoft.com/)
[![Financial Modeling](https://img.shields.io/badge/Standard-FAST%20%7C%20ICAEW-blueviolet.svg)](https://www.fast-standard.org/)

An end-to-end financial data pipeline, SQLite DataOps engine, integrated 3-statement financial model, and Power BI developer setup for **AMOK EQUIPMENT AS** (Org.nr: `911823705`).

---

## Executive Summary

This repository integrates corporate accounting data extraction, production database storage, financial statement modeling, and BI analytics into a governed, auditable decision-support architecture.

### Key Capabilities
1. **Automated API Extraction**: Python pipeline fetching structured financial statements (2023–2025) directly from Brønnøysundregistrenes Regnskapsregister API.
2. **Integrated 3-Statement Excel Model**: Dynamic 8-year model (2023–2030) built in Excel (`models/Amok_Equipment_Integrated_3_Statement_Model.xlsx`) adhering to **FAST Standard** and **ICAEW Financial Modelling Code**. Includes 4 scenario cases, 4-line corkscrew schedules, dynamic sensitivity heatmaps, tornado elasticity, and 12 automated checks.
3. **Institutional SQLite DataOps Database**: Operational SQLite database (`database/amok_equipment.db`) configured with `PRAGMA journal_mode = WAL;`, `foreign_keys = ON;`, star-schema fact tables, analytical views, and cryptographic SHA-256 audit logging.
4. **Power BI Developer Mode (.pbip)**: Modern Power BI Project setup (`powerbi/amok_equipment_project.pbip`) utilizing TMDL semantic model definitions and PBIR report layouts for Git version control.
5. **AI Governance & Risk Framework**: Enterprise AI modeling framework referencing **Federal Reserve SR 11-7**, **EU AI Act**, **DORA**, and Explainable AI (SHAP/LIME).

---

## Repository Architecture

```
Amok Equipment/
├── data/
│   ├── raw/
│   │   └── amok_cash_flow_2024_2025.csv          # Granular indirect cash flow source
│   └── processed/
│       ├── amok_equipment_alle_aar_wide.csv      # Processed wide-format dataset
│       ├── amok_equipment_alle_aar_long.csv      # Normalized star-schema long dataset
│       ├── raw_vendor_data.csv                   # Vendor transaction dataset
│       └── canonical_vendor_name.csv             # Master vendor reference dimension
├── database/
│   ├── amok_equipment.db                         # Production SQLite DB (Amok Financials)
│   └── vendor_finance.db                         # Operational SQLite DB (Vendor Ingestion)
├── docs/
│   ├── ai_governance_ethics_reporting_framework.md # SR 11-7 & EU AI Act governance
│   ├── data_ingestion_sqlite_best_practices.md     # SQLite DataOps engineering guide
│   └── three_statement_modeling_best_practices.md  # FAST / ICAEW Excel modeling guide
├── models/
│   └── Amok_Equipment_Integrated_3_Statement_Model.xlsx # Integrated 3-Statement Model
├── powerbi/
│   ├── amok_equipment_project.pbip              # Power BI Project (Developer Mode)
│   ├── amok_equipment_project.Report/           # Power BI report visual definitions
│   └── amok_equipment_project.SemanticModel/    # Power BI TMDL semantic model
├── src/
│   ├── amok_regnskap_powerbi.py                 # BrReg API financial extractor
│   ├── build_amok_three_statement_model.py      # Excel 3-statement model compiler
│   ├── create_amok_sqlite_db.py                 # SQLite database & views compiler
│   └── sqlite_ingestion_pipeline.py             # Vendor ETL & audit log pipeline
├── .gitignore
├── LES_MEG.txt                                  # Quick-start execution guide (Norwegian)
├── README.md                                    # Repository documentation
└── requirements.txt                             # Python dependencies
```

---

## Quick Start & Execution Guide

### Prerequisites
* Python 3.10+
* Git
* Microsoft Excel 2016+ / Microsoft 365
* Power BI Desktop (for `.pbip` project)

### 1. Install Dependencies
```bash
python -m pip install -r requirements.txt
```

### 2. Extract Live BrReg Accounting Data
Fetches annual financial statements for Org.nr 911823705 and generates normalized wide & long CSV datasets in `data/processed/`:
```bash
python src/amok_regnskap_powerbi.py
```

### 3. Compile Integrated 3-Statement Excel Model
Generates the integrated financial model in `models/Amok_Equipment_Integrated_3_Statement_Model.xlsx`:
```bash
python src/build_amok_three_statement_model.py
```

### 4. Build Production SQLite Database
Builds `database/amok_equipment.db` with relational tables, analytical views, and SHA-256 audit trails:
```bash
python src/create_amok_sqlite_db.py
```

### 5. Run Vendor ETL & Audit Pipeline Test
Executes batch upsert mechanics and financial tie-out verification ($|\text{Source} - \text{DB}| < \$0.01$):
```bash
python src/sqlite_ingestion_pipeline.py
```

### 6. Open Power BI Project
Launch Power BI Desktop and open `powerbi/amok_equipment_project.pbip`.

---

## Financial Highlights & Projection Trajectory

### Historical Actuals (2023–2025)
* **Revenue**: NOK 18.02M (2023) → NOK 20.37M (2024) → NOK 16.23M (2025)
* **Operating Profit (EBIT)**: NOK 2.40M (2023) → NOK 0.75M (2024) → -NOK 1.08M (2025)
* **Net Income**: NOK 3.19M (2023) → NOK 0.63M (2024) → -NOK 0.70M (2025)
* **Ending Cash**: NOK 9.02M (2023) → NOK 8.52M (2024) → NOK 6.91M (2025)

### 2030 Forecast Outcomes by Scenario

| Metric | Base Case (2% Vol, 2% Price) | Bull Case (10% Vol, 3% Price) | Bear Case (-4% Vol, 1% Price) | Severe Stress (-15% Vol, -2% Price) |
| :--- | :--- | :--- | :--- | :--- |
| **Revenue** | NOK 28.47M | NOK 43.89M | NOK 19.98M | NOK 9.69M |
| **EBITDA** | NOK 1.54M | NOK 6.27M | -NOK 1.06M | -NOK 3.25M |
| **EBITDA Margin** | 5.4% | 14.3% | -5.3% | -33.6% |
| **Ending Cash** | NOK 9.08M | NOK 24.32M | NOK 0.38M | NOK 0.00M |
| **Revolver Debt** | NOK 0.00M | NOK 0.00M | NOK 0.00M | NOK 5.00M *(Max Drawn)* |
| **Liquidity Status** | ✅ **PASS** | ✅ **PASS** | ⚠️ **BELOW MIN CASH** | 🚨 **BREACH** |

---

## Core Modeling & Engineering Standards

### 1. FAST & ICAEW Financial Modeling Rules
* **Strict Role Separation**: Dedicated sheets for Inputs (`Assumptions_and_Inputs`), Workings (`Supporting_Schedules`), and Outputs (`Financial_Statements`).
* **Visual Hierarchy**: Inputs styled with light blue fill (`#DCE6F1`) and blue text (`#0000FF`); formulas styled with green text (`#008000`).
* **4-Line Corkscrews**: $\text{Closing}_t = \text{Opening}_t + \text{Additions}_t - \text{Reductions}_t$.
* **Zero Unmanaged Circularity**: Interest expense calculated on opening balances.

### 2. SQLite DataOps & Governance Standards
* **Concurrency & Safety**: `PRAGMA journal_mode = WAL;`, `foreign_keys = ON;`, `busy_timeout = 5000;`.
* **Idempotent Upserts**: `INSERT INTO ... ON CONFLICT DO UPDATE`.
* **Tie-Out Reconciliation**: Automatic verification ensuring $|\text{Source Total} - \text{Ingested Total}| < \$0.01$.
* **Cryptographic Lineage**: SHA-256 hash chaining on all ingestion events recorded in `audit_log`.

---

## License & Attribution

Developed for corporate financial intelligence, BI analytics, and project controlling workflows for **Amok Equipment AS**.  
*Standards referenced: FAST Standard v02c, ICAEW Financial Modelling Code, Federal Reserve SR 11-7, EU AI Act.*
