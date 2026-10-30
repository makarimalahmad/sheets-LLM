import os
import json
import sqlite3
from typing import List, Optional, Dict, Any
from contextlib import contextmanager

DEFAULT_DB_PATH = os.path.join("data", "app.db")

# Buka koneksi ke file database SQLite
@contextmanager
def get_db(db_path: str = DEFAULT_DB_PATH):
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()

# Buat tabel jika belum ada
def init_db(db_path: str = DEFAULT_DB_PATH) -> None:
    with get_db(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS consolidated_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                batch_id TEXT,
                source_sheet TEXT,
                data_json TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS consolidated_batches (
                batch_id TEXT PRIMARY KEY,
                canonical_fields TEXT,
                sheet_mappings TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()

# Simpan hasil gabungan baris multi-sheet beserta pemetaannya
def save_consolidated_records(batch_id: str, records: List[Dict[str, Any]], canonical_fields: Optional[List[str]] = None, sheet_mappings: Optional[Dict[str, Any]] = None, db_path: str = DEFAULT_DB_PATH) -> int:
    init_db(db_path)
    inserted = 0
    with get_db(db_path) as conn:
        cursor = conn.cursor()
        if canonical_fields is not None and sheet_mappings is not None:
            cursor.execute("""
                INSERT OR REPLACE INTO consolidated_batches (batch_id, canonical_fields, sheet_mappings)
                VALUES (?, ?, ?)
            """, (batch_id, json.dumps(canonical_fields), json.dumps(sheet_mappings)))
        for rec in records:
            source = rec.get("_source_sheet", "Unknown")
            cursor.execute("""
                INSERT INTO consolidated_records (batch_id, source_sheet, data_json)
                VALUES (?, ?, ?)
            """, (batch_id, source, json.dumps(rec)))
            inserted += 1
        conn.commit()
    return inserted

# Ambil data gabungan dari batch penggabungan paling terakhir
def get_latest_consolidated_records(limit: int = 100, db_path: str = DEFAULT_DB_PATH) -> Dict[str, Any]:
    init_db(db_path)
    with get_db(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT batch_id FROM consolidated_records ORDER BY id DESC LIMIT 1")
        latest_row = cursor.fetchone()
        if not latest_row:
            return {"records": [], "canonical_fields": [], "sheet_mappings": {}}
        latest_batch = latest_row["batch_id"]

        cursor.execute("SELECT canonical_fields, sheet_mappings FROM consolidated_batches WHERE batch_id = ?", (latest_batch,))
        batch_meta = cursor.fetchone()
        canonical_fields = json.loads(batch_meta["canonical_fields"]) if batch_meta and batch_meta["canonical_fields"] else []
        sheet_mappings = json.loads(batch_meta["sheet_mappings"]) if batch_meta and batch_meta["sheet_mappings"] else {}

        cursor.execute("""
            SELECT id, batch_id, source_sheet, data_json, created_at
            FROM consolidated_records
            WHERE batch_id = ?
            ORDER BY id ASC
            LIMIT ?
        """, (latest_batch, limit))
        rows = cursor.fetchall()
        records = []
        for r in rows:
            item = json.loads(r["data_json"])
            item["_id"] = r["id"]
            item["_batch_id"] = r["batch_id"]
            item["_created_at"] = r["created_at"]
            records.append(item)
        return {
            "records": records,
            "canonical_fields": canonical_fields,
            "sheet_mappings": sheet_mappings
        }

# Kosongkan semua data riwayat di database
def clear_all_records(db_path: str = DEFAULT_DB_PATH) -> None:
    init_db(db_path)
    with get_db(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM consolidated_records")
        cursor.execute("DELETE FROM consolidated_batches")
        conn.commit()
