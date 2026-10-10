"""
AMOK EQUIPMENT AS — Production SQLite Database & Pipeline Setup
===============================================================
Creates amok_equipment.db following institutional SQLite DataOps best practices:
- PRAGMA WAL & foreign keys enabled
- Dimensional tables (company_dim, financial_metrics_dim)
- Normalized fact tables (financial_metrics_long, financial_statements_wide, cash_flow_reconciliation)
- Analytical views (v_pnl_summary, v_balance_sheet_summary, v_key_ratios)
- Atomic upsert mechanics (ON CONFLICT DO UPDATE)
- Financial reconciliation tie-out (|Source - DB| < $0.01)
- Append-only audit logging with SHA-256 hash chaining
"""

import csv
import hashlib
import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "database" / "amok_equipment.db"
WIDE_CSV = BASE_DIR / "data" / "processed" / "amok_equipment_alle_aar_wide.csv"
LONG_CSV = BASE_DIR / "data" / "processed" / "amok_equipment_alle_aar_long.csv"
CASH_FLOW_CSV = BASE_DIR / "data" / "raw" / "amok_cash_flow_2024_2025.csv"



def get_utc_timestamp() -> str:
    return datetime.now(timezone.utc).isoformat()


def init_database(db_path: Path = DB_PATH) -> sqlite3.Connection:
    """Initializes SQLite database with WAL mode and institutional schema."""
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute("PRAGMA busy_timeout = 5000;")

    cursor = conn.cursor()

    # 1. Cryptographic Audit Log
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

    # 2. Company Dimension
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS company_dim (
        organisasjonsnummer TEXT PRIMARY KEY,
        selskap TEXT NOT NULL,
        regnskapstype TEXT NOT NULL DEFAULT 'SELSKAP',
        created_at TEXT NOT NULL
    );
    """)

    # 3. Metric Dimension
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS financial_metrics_dim (
        nokkel_kode TEXT PRIMARY KEY,
        nokkel_navn TEXT NOT NULL,
        kategori TEXT NOT NULL,
        underkategori TEXT NOT NULL,
        enhet TEXT NOT NULL,
        sortering INTEGER NOT NULL
    );
    """)

    # 4. Long Fact Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS financial_metrics_long (
        organisasjonsnummer TEXT NOT NULL,
        regnskapsaar INTEGER NOT NULL,
        nokkel_kode TEXT NOT NULL,
        verdi REAL NOT NULL,
        periode_fra TEXT,
        periode_til TEXT,
        valuta TEXT DEFAULT 'NOK',
        updated_at TEXT NOT NULL,
        PRIMARY KEY (organisasjonsnummer, regnskapsaar, nokkel_kode),
        FOREIGN KEY (organisasjonsnummer) REFERENCES company_dim(organisasjonsnummer),
        FOREIGN KEY (nokkel_kode) REFERENCES financial_metrics_dim(nokkel_kode)
    );
    """)

    # 5. Wide Fact Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS financial_statements_wide (
        organisasjonsnummer TEXT NOT NULL,
        regnskapsaar INTEGER NOT NULL,
        selskap TEXT NOT NULL,
        salgsinntekter REAL,
        driftsinntekter REAL,
        lonnskostnad REAL,
        driftskostnader REAL,
        driftsresultat REAL,
        netto_finans REAL,
        resultat_for_skatt REAL,
        skattekostnad REAL,
        arsresultat REAL,
        anleggsmidler REAL,
        omlopsmidler REAL,
        varer REAL,
        fordringer REAL,
        bank REAL,
        sum_eiendeler REAL,
        egenkapital REAL,
        kortsiktig_gjeld REAL,
        langsiktig_gjeld REAL,
        sum_gjeld REAL,
        driftsmargin_pct REAL,
        resultatmargin_pct REAL,
        egenkapitalandel_pct REAL,
        gjeldsgrad REAL,
        likviditetsgrad_1 REAL,
        likviditetsgrad_2 REAL,
        ROA_pct REAL,
        ROE_pct REAL,
        omsetningsvekst_pct REAL,
        resultatvekst_pct REAL,
        egenkapitalvekst_pct REAL,
        PRIMARY KEY (organisasjonsnummer, regnskapsaar)
    );
    """)

    # 6. Cash Flow Reconciliation Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS cash_flow_reconciliation (
        organisasjonsnummer TEXT NOT NULL DEFAULT '911823705',
        regnskapsaar INTEGER NOT NULL,
        type TEXT NOT NULL,
        navn TEXT NOT NULL,
        belop REAL NOT NULL,
        PRIMARY KEY (organisasjonsnummer, regnskapsaar, navn)
    );
    """)

    # Views
    cursor.execute("""
    CREATE VIEW IF NOT EXISTS v_pnl_summary AS
    SELECT 
        regnskapsaar,
        driftsinntekter AS total_revenue,
        driftskostnader AS operating_expenses,
        lonnskostnad AS personnel_expenses,
        driftsresultat AS operating_profit_ebit,
        netto_finans AS net_financials,
        arsresultat AS net_income
    FROM financial_statements_wide
    ORDER BY regnskapsaar;
    """)

    cursor.execute("""
    CREATE VIEW IF NOT EXISTS v_balance_sheet_summary AS
    SELECT 
        regnskapsaar,
        anleggsmidler AS fixed_assets,
        omlopsmidler AS current_assets,
        sum_eiendeler AS total_assets,
        egenkapital AS equity,
        kortsiktig_gjeld AS short_term_debt,
        langsiktig_gjeld AS long_term_debt,
        sum_gjeld AS total_debt,
        (sum_eiendeler - (sum_gjeld + egenkapital)) AS balance_check_diff
    FROM financial_statements_wide
    ORDER BY regnskapsaar;
    """)

    cursor.execute("""
    CREATE VIEW IF NOT EXISTS v_key_ratios AS
    SELECT 
        regnskapsaar,
        driftsmargin_pct,
        resultatmargin_pct,
        egenkapitalandel_pct,
        gjeldsgrad,
        likviditetsgrad_1,
        likviditetsgrad_2,
        ROA_pct,
        ROE_pct,
        omsetningsvekst_pct
    FROM financial_statements_wide
    ORDER BY regnskapsaar;
    """)

    conn.commit()
    return conn


def get_last_audit_hash(conn: sqlite3.Connection) -> str:
    cursor = conn.cursor()
    cursor.execute("SELECT curr_hash FROM audit_log ORDER BY log_id DESC LIMIT 1;")
    row = cursor.fetchone()
    return row[0] if row else "0" * 64


def log_audit_event(conn: sqlite3.Connection, event_type: str, records_affected: int, total_amount: float, payload: dict) -> None:
    prev_hash = get_last_audit_hash(conn)
    timestamp = get_utc_timestamp()
    payload_str = json.dumps(payload, sort_keys=True)
    hash_data = f"{prev_hash}|{timestamp}|{event_type}|{records_affected}|{total_amount:.2f}|{payload_str}"
    curr_hash = hashlib.sha256(hash_data.encode("utf-8")).hexdigest()

    cursor = conn.cursor()
    cursor.execute("""
    INSERT INTO audit_log (timestamp, event_type, records_affected, total_amount, prev_hash, curr_hash, payload_json)
    VALUES (?, ?, ?, ?, ?, ?, ?);
    """, (timestamp, event_type, records_affected, total_amount, prev_hash, curr_hash, payload_str))
    conn.commit()


def ingest_data(conn: sqlite3.Connection) -> dict:
    cursor = conn.cursor()
    now_str = get_utc_timestamp()

    # 1. Company Dim
    cursor.execute("""
    INSERT INTO company_dim (organisasjonsnummer, selskap, regnskapstype, created_at)
    VALUES ('911823705', 'AMOK EQUIPMENT AS', 'SELSKAP', ?)
    ON CONFLICT(organisasjonsnummer) DO UPDATE SET selskap=EXCLUDED.selskap;
    """, (now_str,))

    # 2. Long Format & Dimension Ingestion
    long_count = 0
    long_total_val = 0.0
    dim_records = {}
    long_rows = []

    with LONG_CSV.open("r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=";")
        for row in reader:
            orgnr = row["organisasjonsnummer"]
            year = int(row["regnskapsaar"])
            kode = row["nokkel_kode"]
            navn = row["nokkel_navn"]
            kategori = row["kategori"]
            underkat = row["underkategori"]
            enhet = row["enhet"]
            sort_order = int(row["sortering"])
            val = float(row["verdi"].replace(",", "."))
            fra = row["periode_fra"]
            til = row["periode_til"]
            valuta = row["valuta"]

            dim_records[kode] = (kode, navn, kategori, underkat, enhet, sort_order)
            long_rows.append((orgnr, year, kode, val, fra, til, valuta))
            long_count += 1
            long_total_val += val

    # Insert parent dimension rows first
    for kode, (k, n, kat, ukat, enh, s) in dim_records.items():
        cursor.execute("""
        INSERT INTO financial_metrics_dim (nokkel_kode, nokkel_navn, kategori, underkategori, enhet, sortering)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(nokkel_kode) DO UPDATE SET
            nokkel_navn=EXCLUDED.nokkel_navn,
            kategori=EXCLUDED.kategori,
            underkategori=EXCLUDED.underkategori,
            enhet=EXCLUDED.enhet,
            sortering=EXCLUDED.sortering;
        """, (k, n, kat, ukat, enh, s))

    # Insert child fact rows after dimension exists
    for orgnr, year, kode, val, fra, til, valuta in long_rows:
        cursor.execute("""
        INSERT INTO financial_metrics_long (
            organisasjonsnummer, regnskapsaar, nokkel_kode, verdi, periode_fra, periode_til, valuta, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(organisasjonsnummer, regnskapsaar, nokkel_kode) DO UPDATE SET
            verdi=EXCLUDED.verdi,
            updated_at=EXCLUDED.updated_at;
        """, (orgnr, year, kode, val, fra, til, valuta, now_str))


    # 3. Wide Format Ingestion
    wide_count = 0
    with WIDE_CSV.open("r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=";")
        for row in reader:
            def safe_f(k):
                v = row.get(k)
                return float(v.replace(",", ".")) if v else None

            cursor.execute("""
            INSERT INTO financial_statements_wide (
                organisasjonsnummer, regnskapsaar, selskap, salgsinntekter, driftsinntekter, lonnskostnad,
                driftskostnader, driftsresultat, netto_finans, resultat_for_skatt, skattekostnad,
                arsresultat, anleggsmidler, omlopsmidler, varer, fordringer, bank, sum_eiendeler,
                egenkapital, kortsiktig_gjeld, langsiktig_gjeld, sum_gjeld, driftsmargin_pct,
                resultatmargin_pct, egenkapitalandel_pct, gjeldsgrad, likviditetsgrad_1, likviditetsgrad_2,
                ROA_pct, ROE_pct, omsetningsvekst_pct, resultatvekst_pct, egenkapitalvekst_pct
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(organisasjonsnummer, regnskapsaar) DO UPDATE SET
                driftsinntekter=EXCLUDED.driftsinntekter,
                arsresultat=EXCLUDED.arsresultat,
                sum_eiendeler=EXCLUDED.sum_eiendeler;
            """, (
                row["organisasjonsnummer"], int(row["regnskapsaar"]), row["selskap"],
                safe_f("salgsinntekter"), safe_f("driftsinntekter"), safe_f("lonnskostnad"),
                safe_f("driftskostnader"), safe_f("driftsresultat"), safe_f("netto_finans"),
                safe_f("resultat_for_skatt"), safe_f("skattekostnad"), safe_f("arsresultat"),
                safe_f("anleggsmidler"), safe_f("omlopsmidler"), safe_f("varer"),
                safe_f("fordringer"), safe_f("bank"), safe_f("sum_eiendeler"),
                safe_f("egenkapital"), safe_f("kortsiktig_gjeld"), safe_f("langsiktig_gjeld"),
                safe_f("sum_gjeld"), safe_f("driftsmargin_pct"), safe_f("resultatmargin_pct"),
                safe_f("egenkapitalandel_pct"), safe_f("gjeldsgrad"), safe_f("likviditetsgrad_1"),
                safe_f("likviditetsgrad_2"), safe_f("ROA_pct"), safe_f("ROE_pct"),
                safe_f("omsetningsvekst_pct"), safe_f("resultatvekst_pct"), safe_f("egenkapitalvekst_pct")
            ))
            wide_count += 1

    # 4. Cash Flow Ingestion
    cf_count = 0
    with CASH_FLOW_CSV.open("r", encoding="cp1252") as f:
        reader = csv.DictReader(f)
        for row in reader:
            cursor.execute("""
            INSERT INTO cash_flow_reconciliation (organisasjonsnummer, regnskapsaar, type, navn, belop)
            VALUES ('911823705', ?, ?, ?, ?)
            ON CONFLICT(organisasjonsnummer, regnskapsaar, navn) DO UPDATE SET belop=EXCLUDED.belop;
            """, (int(row["Kontantstrøm År"]), row["Type"], row["Navn"], float(row["Beløp"])))
            cf_count += 1

    conn.commit()

    # Reconcile counts & check math tie-outs
    cursor.execute("SELECT COUNT(*), SUM(verdi) FROM financial_metrics_long")
    db_long_count, db_long_sum = cursor.fetchone()

    tie_out_passed = (db_long_count == long_count) and (abs(db_long_sum - long_total_val) < 0.01)

    summary_payload = {
        "long_records_ingested": db_long_count,
        "long_sum_value": db_long_sum,
        "wide_years_ingested": wide_count,
        "cash_flow_records_ingested": cf_count,
        "tie_out_passed": tie_out_passed
    }

    log_audit_event(conn, "DATABASE_INGESTION_COMPLETE", db_long_count + wide_count + cf_count, db_long_sum, summary_payload)
    return summary_payload


if __name__ == "__main__":
    print("=== CREATING AMOK EQUIPMENT SQLITE DATABASE ===")
    conn = init_database()
    summary = ingest_data(conn)
    conn.close()

    print(f"* Target Database:       {DB_PATH}")
    print(f"* Long Fact Records:     {summary['long_records_ingested']} (Sum: NOK {summary['long_sum_value']:,.2f})")
    print(f"* Wide Financial Years:  {summary['wide_years_ingested']}")
    print(f"* Cash Flow Statements:  {summary['cash_flow_records_ingested']}")
    print(f"* Tie-Out Assertion:     {'PASSED [OK]' if summary['tie_out_passed'] else 'FAILED [ERR]'}")
