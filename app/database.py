import os
import json
import sqlite3
from typing import List, Optional, Dict, Any
from contextlib import contextmanager

from app.models import StandardProductItem, ProductAnalytics

DEFAULT_DB_PATH = os.path.join("data", "app.db")

@contextmanager
def get_db(db_path: str = DEFAULT_DB_PATH):
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()

def init_db(db_path: str = DEFAULT_DB_PATH) -> None:
    with get_db(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source TEXT,
                eup_inc_vat REAL,
                eup_ex_vat REAL,
                payment_rate REAL,
                payment_cost REAL,
                unit_price REAL,
                cogs REAL,
                margin REAL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS consolidated_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                batch_id TEXT,
                source_sheet TEXT,
                data_json TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()

def save_products(items: List[StandardProductItem], db_path: str = DEFAULT_DB_PATH) -> int:
    init_db(db_path)
    inserted_count = 0
    with get_db(db_path) as conn:
        cursor = conn.cursor()
        for item in items:
            cursor.execute("""
                INSERT INTO products (
                    source,
                    eup_inc_vat,
                    eup_ex_vat,
                    payment_rate,
                    payment_cost,
                    unit_price,
                    cogs,
                    margin
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                item.source,
                item.eup_inc_vat,
                item.eup_ex_vat,
                item.payment_rate,
                item.payment_cost,
                item.unit_price,
                item.cogs,
                item.margin
            ))
            inserted_count += 1
        conn.commit()
    return inserted_count

def save_consolidated_records(batch_id: str, records: List[Dict[str, Any]], db_path: str = DEFAULT_DB_PATH) -> int:
    init_db(db_path)
    inserted = 0
    with get_db(db_path) as conn:
        cursor = conn.cursor()
        for rec in records:
            source = rec.get("_source_sheet", "Unknown")
            cursor.execute("""
                INSERT INTO consolidated_records (batch_id, source_sheet, data_json)
                VALUES (?, ?, ?)
            """, (batch_id, source, json.dumps(rec)))
            inserted += 1
        conn.commit()
    return inserted

def get_latest_consolidated_records(limit: int = 100, db_path: str = DEFAULT_DB_PATH) -> List[Dict[str, Any]]:
    init_db(db_path)
    with get_db(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, batch_id, source_sheet, data_json, created_at
            FROM consolidated_records
            ORDER BY id DESC
            LIMIT ?
        """, (limit,))
        rows = cursor.fetchall()
        result = []
        for r in rows:
            item = json.loads(r["data_json"])
            item["_id"] = r["id"]
            item["_batch_id"] = r["batch_id"]
            item["_created_at"] = r["created_at"]
            result.append(item)
        return result

def get_all_products(limit: int = 100, db_path: str = DEFAULT_DB_PATH) -> List[Dict[str, Any]]:
    init_db(db_path)
    with get_db(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT * FROM products ORDER BY id DESC LIMIT ?
        """, (limit,))
        rows = cursor.fetchall()
        return [dict(row) for row in rows]

def get_highest_eup(db_path: str = DEFAULT_DB_PATH) -> Optional[Dict[str, Any]]:
    init_db(db_path)
    with get_db(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT * FROM products ORDER BY eup_inc_vat DESC LIMIT 1
        """)
        row = cursor.fetchone()
        return dict(row) if row else None

def get_analytics(db_path: str = DEFAULT_DB_PATH) -> ProductAnalytics:
    init_db(db_path)
    with get_db(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT 
                COUNT(*) as total_products,
                MAX(eup_inc_vat) as max_eup,
                MIN(eup_inc_vat) as min_eup,
                AVG(margin) as avg_margin
            FROM products
        """)
        stats = cursor.fetchone()
        
    highest_product = get_highest_eup(db_path)

    return ProductAnalytics(
        total_products=stats["total_products"] or 0,
        highest_eup_inc_vat=stats["max_eup"],
        lowest_eup_inc_vat=stats["min_eup"],
        average_margin=stats["avg_margin"],
        top_product_by_eup=highest_product
    )
