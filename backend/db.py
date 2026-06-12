import os
import json
import sqlite3
from typing import Any, Dict, List, Tuple

DB_PATH = os.getenv("DB_PATH", "/home/data/reports.db")


def get_conn() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(os.path.abspath(DB_PATH)), exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    """Create the reports table if it doesn't exist."""
    with get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS reports (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                data_json   TEXT    NOT NULL,
                uploaded_at TEXT    NOT NULL
            )
        """)


def insert_report(flat_row: Dict[str, Any]) -> int:
    """Insert a flat row dict as JSON. Returns the new row id."""
    uploaded_at = flat_row.get("uploaded_at", "")
    with get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO reports (data_json, uploaded_at) VALUES (?, ?)",
            (json.dumps(flat_row, ensure_ascii=False, default=str), uploaded_at),
        )
        return cur.lastrowid


def get_all_reports() -> Tuple[List[str], List[Dict]]:
    """
    Returns (headers, rows) where each row is a dict with 'row_idx' and 'cells'.
    Columns are derived from the union of all keys across all rows.
    """
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT id, data_json FROM reports ORDER BY id DESC"
        ).fetchall()

    if not rows:
        return [], []

    parsed = [(r["id"], json.loads(r["data_json"])) for r in rows]

    # Build ordered headers: id first, then union of all keys
    all_keys: list = ["id"]
    seen = {"id"}
    for _, data in parsed:
        for k in data:
            if k not in seen:
                all_keys.append(k)
                seen.add(k)

    result_rows = []
    for row_id, data in parsed:
        cells = [row_id] + [data.get(k) for k in all_keys[1:]]
        result_rows.append({"row_idx": row_id, "cells": cells})

    return all_keys, result_rows


def delete_report(row_id: int) -> bool:
    """Delete by id. Returns True if a row was deleted."""
    with get_conn() as conn:
        cur = conn.execute("DELETE FROM reports WHERE id = ?", (row_id,))
        return cur.rowcount > 0


def export_all_flat_rows() -> Tuple[List[str], List[List[Any]]]:
    """
    Returns (headers, list_of_value_lists) for Excel export.
    """
    headers, rows = get_all_reports()
    value_lists = [row["cells"] for row in rows]
    return headers, value_lists
