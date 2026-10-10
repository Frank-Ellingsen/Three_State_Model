# Best Practice Guide: Data Gathering, Ingestion, SQLite Storage, and Delta Management

This engineering guide defines institutional best practices for gathering, profiling, ingesting, and storing structured and unstructured data using **SQLite** as an operational database. Grounded in enterprise data governance principles, financial modeling standards (**ICAEW**, **FAST Standard**, **SSRB**), and modern DataOps architectures, this document outlines an auditable, deterministic workflow.

---

## Executive Summary & Core Engineering Principles

1. **Profile Before Fixing**: Measure data quality issues (blanks, duplicates, out-of-range values) before applying transformations.
2. **Separation of Ingestion & Transformation**: Stage raw inputs in immutable landing tables before transforming data into production canonical schemas.
3. **Validation by Comparison**: Every ingestion step must verify row counts and financial totals against raw source control figures ($|\text{Ingested Total} - \text{Source Total}| < \$0.01$).
4. **Idempotent Upserts**: Use deterministic primary keys and `ON CONFLICT` clauses so re-running pipelines never creates duplicate records or corrupted state.
5. **Immutable Audit Trails**: Record row-level changes, source lineage, and modification timestamps with cryptographic checksums for complete audit defensibility.

---

## 1. Data Gathering & Pre-Ingestion Profiling

### A. Heterogeneous Source Acquisition
Data pipelines ingest data across multiple modalities:
* **Structured Data**: General Ledger actuals, ERP databases (Unit4 UBW, SAP, Oracle), local SQLite databases (`projects.db`), CSV/Parquet files.
* **Semi-Structured Data**: Rest APIs (JSON payloads via Model Context Protocol / MCP), web hooks, XML disclosures.
* **Unstructured Text**: Vendor contracts, invoice PDFs, SEC EDGAR filings (parsed via OCR / Layout-aware extractors like FinBERT).

### B. Pre-Ingestion Data Profiling (The 4 Quality Tests)
Before inserting data into SQLite, evaluate the dataset against the four foundational quality criteria:
* **Completeness**: Assert zero missing required fields (e.g., `Transaction_ID`, `Posting_Date`, `Amount`, `Cost_Center`).
* **Consistency**: Ensure uniform formatting (e.g., dates in ISO-8601 `YYYY-MM-DD`, normalized currency codes `USD`, `NOK`, uppercase state/country codes).
* **Accuracy**: Verify numerical values fall within realistic business bounds (e.g., standard deviation outlier checks $> 3\sigma$, non-negative amounts where prohibited).
* **Uniqueness**: Detect duplicate primary key signatures or content hashes to prevent inflated totals.

```python
import pandas as pd

def profile_raw_dataset(df: pd.DataFrame, primary_key: str) -> dict:
    """Profiles incoming dataset health prior to ingestion."""
    return {
        "total_records": len(df),
        "null_summary": df.isnull().sum().to_dict(),
        "duplicate_keys": df.duplicated(subset=[primary_key]).sum(),
        "date_range": (df["Invoice_Date"].min(), df["Invoice_Date"].max()),
        "total_numeric_volume": float(df["Invoice_Amount"].sum())
    }
```

---

## 2. SQLite Database Schema Design & Performance Setup

### A. Operational Configuration (SQLite PRAGMAs)
SQLite requires optimal operational settings for multi-threaded performance, data safety, and foreign key enforcement:

```sql
-- Enable Write-Ahead Logging (WAL) for concurrent read/write throughput
PRAGMA journal_mode = WAL;

-- Enforce relational foreign key constraints
PRAGMA foreign_keys = ON;

-- Set busy timeout to handle lock contention gracefully
PRAGMA busy_timeout = 5000;

-- Optimize memory cache size (e.g., 64MB)
PRAGMA cache_size = -64000;

-- Enforce UTF-8 text encoding
PRAGMA encoding = "UTF-8";
```

### B. Canonical Relational Schema Architecture
Decouple **Raw Staging** from **Canonical Production Entities** to insulate analytics from shifting source layouts.

```sql
-- 1. Immutable Raw Staging Table
CREATE TABLE IF NOT EXISTS staging_vendor_invoices (
    staging_id INTEGER PRIMARY KEY AUTOINCREMENT,
    raw_vendor_name TEXT NOT NULL,
    invoice_number TEXT,
    invoice_date TEXT NOT NULL,
    invoice_amount REAL NOT NULL,
    department_code TEXT,
    source_file TEXT NOT NULL,
    ingested_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 2. Master Canonical Vendor Dimension
CREATE TABLE IF NOT EXISTS dim_canonical_vendors (
    vendor_id INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_name TEXT UNIQUE NOT NULL,
    acronym_alias TEXT,
    is_active INTEGER DEFAULT 1 CHECK (is_active IN (0, 1)),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 3. Core Transactions Fact Table
CREATE TABLE IF NOT EXISTS fact_vendor_transactions (
    transaction_id TEXT PRIMARY KEY,
    vendor_id INTEGER NOT NULL,
    invoice_date TEXT NOT NULL, -- Stored as ISO-8601 YYYY-MM-DD
    invoice_amount REAL NOT NULL CHECK (invoice_amount >= 0),
    department TEXT NOT NULL,
    confidence_score REAL DEFAULT 1.0,
    needs_human_review INTEGER DEFAULT 0 CHECK (needs_human_review IN (0, 1)),
    audit_reasoning TEXT,
    content_hash TEXT UNIQUE NOT NULL, -- SHA-256 hash for deduplication
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (vendor_id) REFERENCES dim_canonical_vendors(vendor_id)
);

-- 4. Immutable Audit Trail Table
CREATE TABLE IF NOT EXISTS audit_log (
    log_id INTEGER PRIMARY KEY AUTOINCREMENT,
    transaction_id TEXT NOT NULL,
    action_type TEXT NOT NULL, -- 'INSERT', 'UPDATE', 'UPSERT', 'QUARANTINE'
    old_values JSON,
    new_values JSON NOT NULL,
    performed_by TEXT NOT NULL,
    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    record_hash TEXT NOT NULL
);

-- Indexes for Query Acceleration
CREATE INDEX IF NOT EXISTS idx_fact_invoice_date ON fact_vendor_transactions(invoice_date);
CREATE INDEX IF NOT EXISTS idx_fact_vendor_id ON fact_vendor_transactions(vendor_id);
```

---

## 3. Data Ingestion Architecture & Execution

### A. Transactional Batch Ingestion
Execute ingestion within explicit SQLite transaction blocks (`BEGIN TRANSACTION ... COMMIT`) using parameter binding to prevent SQL injection and achieve 10,000+ inserts/second.

```python
import sqlite3
import hashlib
import json

def generate_record_hash(row: dict) -> str:
    """Computes SHA-256 fingerprint of key business attributes."""
    payload = f"{row['Transaction_ID']}|{row['Invoice_Date']}|{row['Invoice_Amount']}|{row['Canonical_Vendor_Name']}"
    return hashlib.sha256(payload.encode('utf-8')).hexdigest()

def batch_ingest_transactions(db_path: str, records: list):
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON;")
    cursor = conn.cursor()
    
    try:
        cursor.execute("BEGIN TRANSACTION;")
        for rec in records:
            # 1. Resolve or create vendor ID
            cursor.execute(
                "INSERT INTO dim_canonical_vendors (canonical_name) VALUES (?) "
                "ON CONFLICT(canonical_name) DO UPDATE SET updated_at=CURRENT_TIMESTAMP "
                "RETURNING vendor_id;",
                (rec['Canonical_Vendor_Name'],)
            )
            vendor_id = cursor.fetchone()[0]
            
            # 2. Compute checksum hash
            rec_hash = generate_record_hash(rec)
            
            # 3. Upsert Fact Transaction
            cursor.execute("""
                INSERT INTO fact_vendor_transactions (
                    transaction_id, vendor_id, invoice_date, invoice_amount, 
                    department, confidence_score, needs_human_review, 
                    audit_reasoning, content_hash
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(transaction_id) DO UPDATE SET
                    vendor_id = EXCLUDED.vendor_id,
                    invoice_date = EXCLUDED.invoice_date,
                    invoice_amount = EXCLUDED.invoice_amount,
                    department = EXCLUDED.department,
                    confidence_score = EXCLUDED.confidence_score,
                    needs_human_review = EXCLUDED.needs_human_review,
                    audit_reasoning = EXCLUDED.audit_reasoning,
                    content_hash = EXCLUDED.content_hash,
                    updated_at = CURRENT_TIMESTAMP;
            """, (
                rec['Transaction_ID'], vendor_id, rec['Invoice_Date'], 
                rec['Invoice_Amount'], rec['Department'], rec['Confidence_Score'],
                1 if rec['Needs_Human_Review'] == 'YES' else 0,
                rec['Audit_Reasoning'], rec_hash
            ))
            
            # 4. Record Audit Log Entry
            cursor.execute("""
                INSERT INTO audit_log (transaction_id, action_type, new_values, performed_by, record_hash)
                VALUES (?, 'UPSERT', ?, 'DATA_PIPELINE_SERVICE', ?);
            """, (rec['Transaction_ID'], json.dumps(rec), rec_hash))
            
        conn.commit()
        print(f"Successfully processed {len(records)} records.")
    except Exception as e:
        conn.rollback()
        print(f"Ingestion failed! Transaction rolled back. Error: {e}")
        raise e
    finally:
        conn.close()
"""
```

---

## 4. Delta Management: Adding, Updating & Upserting New Data

When new transaction files arrive (e.g., daily incremental vendor feeds), pipelines must handle inserts, updates, and exceptions deterministically.

### A. Upsert Pattern (`ON CONFLICT`)
SQLite 3.24+ supports native `INSERT INTO ... ON CONFLICT(key) DO UPDATE SET ...` syntax:

```sql
INSERT INTO fact_vendor_transactions (
    transaction_id, vendor_id, invoice_date, invoice_amount, department, updated_at
) VALUES ('TXN-1004', 3, '2024-01-09', 210000.0, 'Risk Advisory', CURRENT_TIMESTAMP)
ON CONFLICT(transaction_id) DO UPDATE SET
    invoice_amount = EXCLUDED.invoice_amount,
    department = EXCLUDED.department,
    updated_at = CURRENT_TIMESTAMP
WHERE EXCLUDED.invoice_amount != fact_vendor_transactions.invoice_amount;
```

### B. Handling Schema Evolution & Migrations
Track database schema modifications in a dedicated migration ledger to allow seamless additions of new columns without breaking historical records.

```sql
CREATE TABLE IF NOT EXISTS schema_migrations (
    version INTEGER PRIMARY KEY,
    description TEXT NOT NULL,
    applied_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Example Migration: Adding PO_Number column safely
ALTER TABLE fact_vendor_transactions ADD COLUMN po_number TEXT DEFAULT NULL;
INSERT INTO schema_migrations (version, description) VALUES (2, 'Add po_number to fact_vendor_transactions');
```

### C. Human-in-the-Loop (HITL) Exception Quarantining
Low-confidence matches (e.g., confidence score $< 0.82$) or unmapped vendors must be isolated into a quarantine view for human review before updating financial control totals.

```sql
CREATE VIEW IF NOT EXISTS v_quarantine_exceptions AS
SELECT 
    t.transaction_id,
    v.canonical_name,
    t.invoice_amount,
    t.confidence_score,
    t.audit_reasoning
FROM fact_vendor_transactions t
JOIN dim_canonical_vendors v ON t.vendor_id = v.vendor_id
WHERE t.needs_human_review = 1 OR v.canonical_name LIKE 'UNMAPPED%';
```

---

## 5. Auditability, Checksums & Post-Ingestion Reconciliation

### Post-Ingestion Verification Checklist
Run automated assertions immediately after each batch commit:

1. **Row Count Reconciliation**: Assert that count of raw input rows matches staged + processed rows.
   $$\text{Count}_{\text{Staged}} = \text{Count}_{\text{Inserted}} + \text{Count}_{\text{Updated}}$$
2. **Financial Spend Tie-Out**: Reconcile sum of raw invoice amounts against database totals within $\$0.01$ tolerance.
   $$|\sum \text{Invoice\_Amount}_{\text{Raw}} - \sum \text{Invoice\_Amount}_{\text{DB}}| < 0.01$$
3. **Null & Constraint Inspection**: Confirm zero null values in mandatory foreign keys and non-negative amounts.

---

## 6. Summary Workflow Matrix

| Pipeline Stage | Primary SQLite Feature / Python Module | Audit & Governance Objective |
| :--- | :--- | :--- |
| **1. Data Gathering** | `pandas`, `requests`, `PyPDF` | Multi-source extraction, metadata tagging |
| **2. Pre-Cleaning & Profiling** | `df.isnull()`, `difflib` similarity | Quantify missing values & duplicate signatures |
| **3. Database Setup** | SQLite `PRAGMA WAL`, `FOREIGN KEYS = ON` | Concurrency, relational integrity, zero corruption |
| **4. Ingestion Execution** | `BEGIN TRANSACTION`, `executemany` | Atomic execution, rollbacks on failure |
| **5. Incremental Upserts** | `INSERT INTO ... ON CONFLICT` | Idempotent updates, zero duplicate rows |
| **6. Verification & Tie-Out** | SQL `SUM()`, `COUNT()`, Python `assert` | Math tie-out ($< \$0.01$), 100% row preservation |
| **7. Compliance Logging** | `audit_log` table with SHA-256 hashes | Immutable record of who, what, and when |

---
*Generated by Gemini Notebook | Project Finance AI Architecture & FinOps*
