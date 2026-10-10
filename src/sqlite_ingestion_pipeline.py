"""
Production-Grade SQLite Ingestion & Delta Management Pipeline
==============================================================
Demonstrates:
- Parameterized SQLite schema with strict PRAGMA WAL & CHECK constraints
- Fuzzy & exact vendor mapping from master canonical reference
- Atomic batch upsert mechanics (ON CONFLICT DO UPDATE)
- Automated financial reconciliation tie-outs (|Source - Ingested| < $0.01)
- Append-only audit logging with SHA-256 cryptographic hash chaining
"""

import sqlite3
import json
import hashlib
import re
from datetime import datetime
import pandas as pd
from difflib import SequenceMatcher

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_DB_PATH = BASE_DIR / "database" / "vendor_finance.db"


def init_database(db_path=DEFAULT_DB_PATH):
    """Initializes SQLite database with production PRAGMAs and constrained tables."""
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA foreign_keys = ON;")
    
    cursor = conn.cursor()
    
    # 1. Cryptographic Audit Log Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS audit_log (
        log_id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp TEXT NOT NULL,
        event_type TEXT NOT NULL,
        records_affected INTEGER NOT NULL,
        total_amount REAL NOT NULL,
        prev_hash TEXT NOT NULL,
        curr_hash TEXT NOT NULL,
        payload_json TEXT NOT NULL
    );
    """)
    
    # 2. Canonical Vendor Master Reference
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS canonical_vendors (
        canonical_id INTEGER PRIMARY KEY AUTOINCREMENT,
        canonical_name TEXT UNIQUE NOT NULL,
        category TEXT NOT NULL,
        approved_since INTEGER NOT NULL
    );
    """)
    
    # 3. Main Vendor Transactions Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS vendor_transactions (
        transaction_id TEXT PRIMARY KEY,
        raw_vendor_name TEXT NOT NULL,
        canonical_vendor_name TEXT NOT NULL,
        department TEXT NOT NULL,
        invoice_amount REAL NOT NULL CHECK (invoice_amount >= 0),
        invoice_date TEXT NOT NULL,
        confidence_score REAL NOT NULL CHECK (confidence_score BETWEEN 0.0 AND 1.0),
        needs_human_review TEXT NOT NULL CHECK (needs_human_review IN ('YES', 'NO')),
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    );
    """)
    
    # 4. Quarantined Exceptions Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS transaction_exceptions (
        transaction_id TEXT PRIMARY KEY,
        raw_vendor_name TEXT NOT NULL,
        invoice_amount REAL NOT NULL,
        invoice_date TEXT NOT NULL,
        reason TEXT NOT NULL,
        flagged_at TEXT NOT NULL
    );
    """)
    
    conn.commit()
    return conn

def get_last_audit_hash(conn):
    cursor = conn.cursor()
    cursor.execute("SELECT curr_hash FROM audit_log ORDER BY log_id DESC LIMIT 1;")
    row = cursor.fetchone()
    return row[0] if row else "0" * 64

def log_audit_event(conn, event_type, records_affected, total_amount, payload):
    prev_hash = get_last_audit_hash(conn)
    timestamp = datetime.utcnow().isoformat() + "Z"
    payload_str = json.dumps(payload, sort_keys=True)
    
    hash_data = f"{prev_hash}|{timestamp}|{event_type}|{records_affected}|{total_amount:.2f}|{payload_str}"
    curr_hash = hashlib.sha256(hash_data.encode('utf-8')).hexdigest()
    
    cursor = conn.cursor()
    cursor.execute("""
    INSERT INTO audit_log (timestamp, event_type, records_affected, total_amount, prev_hash, curr_hash, payload_json)
    VALUES (?, ?, ?, ?, ?, ?, ?);
    """, (timestamp, event_type, records_affected, total_amount, prev_hash, curr_hash, payload_str))
    conn.commit()

def normalize_string(text):
    text = text.lower()
    text = re.sub(r'[^a-z0-9\s]', '', text)
    tokens = [t for t in text.split() if t not in ['inc', 'corp', 'llc', 'com', 'ltd']]
    return " ".join(tokens)

def match_vendor(raw_name, canonical_list):
    clean_raw = normalize_string(raw_name)
    
    # Alias map
    aliases = {
        'msft': 'Microsoft',
        'iso': 'Insurance Services Office (ISO)',
        'i.s.o.': 'Insurance Services Office (ISO)',
    }
    if raw_name.lower().strip() in aliases:
        return aliases[raw_name.lower().strip()], 1.0
        
    best_match = "UNMAPPED - NEW VENDOR"
    best_score = 0.0
    
    for c_name in canonical_list:
        clean_c = normalize_string(c_name)
        ratio = SequenceMatcher(None, clean_raw, clean_c).ratio()
        if ratio > best_score:
            best_score = ratio
            best_match = c_name
            
    if best_score >= 0.60:
        return best_match, round(best_score, 2)
    else:
        return "UNMAPPED - NEW VENDOR", round(best_score, 2)

def execute_ingestion_pipeline(raw_csv_path, canonical_csv_path, db_path=DEFAULT_DB_PATH):
    conn = init_database(db_path)
    
    # 1. Ingest Canonical Reference
    df_canonical = pd.read_csv(canonical_csv_path)
    cursor = conn.cursor()
    for _, row in df_canonical.iterrows():
        cursor.execute("""
        INSERT OR IGNORE INTO canonical_vendors (canonical_name, category, approved_since)
        VALUES (?, ?, ?);
        """, (row['Canonical Vendor Name'].strip(), row['Category'].strip(), int(row['Approved Since'])))
    conn.commit()
    
    canonical_names = df_canonical['Canonical Vendor Name'].str.strip().tolist()
    
    # 2. Ingest Raw Vendor Data
    df_raw = pd.read_csv(raw_csv_path)
    source_count = len(df_raw)
    source_sum = float(df_raw['Invoice Amount ($)'].sum())
    
    now_str = datetime.utcnow().isoformat() + "Z"
    valid_records = []
    exception_records = []
    
    for _, row in df_raw.iterrows():
        txn_id = str(row['Transaction ID']).strip()
        raw_vendor = str(row['Vendor Name (Raw)']).strip()
        dept = str(row['Department']).strip()
        amount = float(row['Invoice Amount ($)'])
        inv_date = str(row['Invoice Date']).strip()
        
        c_match, conf_score = match_vendor(raw_vendor, canonical_names)
        
        needs_review = "NO" if conf_score >= 0.85 and c_match != "UNMAPPED - NEW VENDOR" else "YES"
        
        if c_match == "UNMAPPED - NEW VENDOR":
            exception_records.append((txn_id, raw_vendor, amount, inv_date, "Unmapped Canonical Vendor", now_str))
        elif amount < 0:
            exception_records.append((txn_id, raw_vendor, amount, inv_date, "Negative Invoice Amount", now_str))
        else:
            valid_records.append((txn_id, raw_vendor, c_match, dept, amount, inv_date, conf_score, needs_review, now_str, now_str))
            
    # 3. Execute Batch Upsert
    upsert_sql = """
    INSERT INTO vendor_transactions (
        transaction_id, raw_vendor_name, canonical_vendor_name, department, 
        invoice_amount, invoice_date, confidence_score, needs_human_review, created_at, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(transaction_id) DO UPDATE SET
        raw_vendor_name = excluded.raw_vendor_name,
        canonical_vendor_name = excluded.canonical_vendor_name,
        department = excluded.department,
        invoice_amount = excluded.invoice_amount,
        invoice_date = excluded.invoice_date,
        confidence_score = excluded.confidence_score,
        needs_human_review = excluded.needs_human_review,
        updated_at = excluded.updated_at;
    """
    cursor.executemany(upsert_sql, valid_records)
    
    exception_sql = """
    INSERT INTO transaction_exceptions (transaction_id, raw_vendor_name, invoice_amount, invoice_date, reason, flagged_at)
    VALUES (?, ?, ?, ?, ?, ?)
    ON CONFLICT(transaction_id) DO UPDATE SET
        invoice_amount = excluded.invoice_amount,
        reason = excluded.reason,
        flagged_at = excluded.flagged_at;
    """
    cursor.executemany(exception_sql, exception_records)
    conn.commit()
    
    # 4. Reconcile
    cursor.execute("SELECT COUNT(*), COALESCE(SUM(invoice_amount), 0.0) FROM vendor_transactions")
    valid_count, valid_sum = cursor.fetchone()
    
    cursor.execute("SELECT COUNT(*), COALESCE(SUM(invoice_amount), 0.0) FROM transaction_exceptions")
    exc_count, exc_sum = cursor.fetchone()
    
    total_ingested_count = valid_count + exc_count
    total_ingested_sum = valid_sum + exc_sum
    
    sum_delta = abs(source_sum - total_ingested_sum)
    reconciliation_passed = (source_count == total_ingested_count) and (sum_delta < 0.01)
    
    audit_payload = {
        "source_count": source_count,
        "source_sum": source_sum,
        "valid_count": valid_count,
        "valid_sum": valid_sum,
        "exception_count": exc_count,
        "exception_sum": exc_sum,
        "reconciliation_passed": reconciliation_passed,
        "sum_delta": sum_delta
    }
    
    log_audit_event(conn, "BATCH_UPSERT_INGESTION", total_ingested_count, total_ingested_sum, audit_payload)
    conn.close()
    
    return audit_payload

def create_sample_files_if_missing(raw_csv_path, canonical_csv_path):
    raw_path = Path(raw_csv_path)
    canonical_path = Path(canonical_csv_path)
    
    if not canonical_path.exists():
        canonical_path.parent.mkdir(parents=True, exist_ok=True)
        canonical_df = pd.DataFrame([
            {"Canonical Vendor Name": "Microsoft", "Category": "Software & Cloud", "Approved Since": 2018},
            {"Canonical Vendor Name": "Insurance Services Office (ISO)", "Category": "Risk & Compliance", "Approved Since": 2015},
            {"Canonical Vendor Name": "Salesforce Inc", "Category": "CRM Software", "Approved Since": 2020},
            {"Canonical Vendor Name": "Amazon Web Services", "Category": "Cloud Infrastructure", "Approved Since": 2017}
        ])
        canonical_df.to_csv(canonical_path, index=False)
        print(f"Created sample canonical vendor reference: {canonical_path}")
        
    if not raw_path.exists():
        raw_path.parent.mkdir(parents=True, exist_ok=True)
        raw_df = pd.DataFrame([
            {"Transaction ID": "TXN-101", "Vendor Name (Raw)": "MSFT", "Department": "IT Ops", "Invoice Amount ($)": 1250.00, "Invoice Date": "2026-09-15"},
            {"Transaction ID": "TXN-102", "Vendor Name (Raw)": "Microsoft Corp.", "Department": "Software", "Invoice Amount ($)": 4800.50, "Invoice Date": "2026-09-18"},
            {"Transaction ID": "TXN-103", "Vendor Name (Raw)": "I.S.O.", "Department": "Risk", "Invoice Amount ($)": 9500.00, "Invoice Date": "2026-09-20"},
            {"Transaction ID": "TXN-104", "Vendor Name (Raw)": "Salesforce.com", "Department": "Sales", "Invoice Amount ($)": 3200.00, "Invoice Date": "2026-09-22"},
            {"Transaction ID": "TXN-105", "Vendor Name (Raw)": "Unknown Vendor Inc", "Department": "General", "Invoice Amount ($)": 750.00, "Invoice Date": "2026-09-25"}
        ])
        raw_df.to_csv(raw_path, index=False)
        print(f"Created sample raw vendor dataset: {raw_path}")

if __name__ == "__main__":
    raw_csv = BASE_DIR / "data" / "processed" / "raw_vendor_data.csv"
    canonical_csv = BASE_DIR / "data" / "processed" / "canonical_vendor_name.csv"
    db_file = DEFAULT_DB_PATH

    
    create_sample_files_if_missing(raw_csv, canonical_csv)
    
    print("=== EXECUTING SQLITE DATA INGESTION & UPSERT PIPELINE ===")
    summary = execute_ingestion_pipeline(raw_csv, canonical_csv, db_file)
    
    print(f"* Source Records Read:  {summary['source_count']}")
    print(f"* Source Total Spend:   ${summary['source_sum']:,.2f}")
    print(f"* Ingested Valid Rows:  {summary['valid_count']} (${summary['valid_sum']:,.2f})")
    print(f"* Quarantined Rows:    {summary['exception_count']} (${summary['exception_sum']:,.2f})")
    print(f"* Total Ingested Spend: ${summary['valid_sum'] + summary['exception_sum']:,.2f}")
    print(f"* Reconciled Tie-Out:  {'PASSED [OK] (Tolerance < $0.01)' if summary['reconciliation_passed'] else 'FAILED [ERR]'}")


