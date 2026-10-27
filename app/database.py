import os
import sqlite3
from typing import List, Optional, Dict, Any
from app.models import StandardProductItem, ProductAnalytics

from contextlib import contextmanager

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
