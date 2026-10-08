from __future__ import annotations

import secrets
import sqlite3
from typing import Any, Dict, Iterable, Optional

from outlook_web.db import get_db


def generate_secret() -> str:
    return f"mail_{secrets.token_urlsafe(24)}"


def create_for_account(db: sqlite3.Connection, account_id: int) -> None:
    db.execute(
        "INSERT INTO customer_mail_credentials (account_id, secret) VALUES (?, ?)",
        (account_id, generate_secret()),
    )


def backfill_credentials(db: sqlite3.Connection, *, migrate_legacy: bool = False) -> None:
    """Preserve secrets from the provisional accounts-column implementation."""
    columns = {row["name"] for row in db.execute("PRAGMA table_info(accounts)")} if migrate_legacy else set()
    if migrate_legacy and "customer_secret" in columns:
        enabled = "customer_secret_enabled" if "customer_secret_enabled" in columns else "1"
        expires = "customer_secret_expires_at" if "customer_secret_expires_at" in columns else "NULL"
        db.execute(
            f"""
            INSERT OR IGNORE INTO customer_mail_credentials (account_id, secret, enabled, expires_at)
            SELECT id, customer_secret, {enabled}, {expires}
            FROM accounts
            WHERE customer_secret IS NOT NULL AND TRIM(customer_secret) <> ''
            """
        )
    for row in db.execute(
        """
        SELECT a.id FROM accounts a
        LEFT JOIN customer_mail_credentials c ON c.account_id = a.id
        WHERE c.account_id IS NULL
        """
    ).fetchall():
        create_for_account(db, row["id"])


def by_account_ids(db: sqlite3.Connection, account_ids: Iterable[int]) -> Dict[int, Dict[str, Any]]:
    ids = list(account_ids)
    if not ids:
        return {}
    placeholders = ",".join("?" for _ in ids)
    rows = db.execute(
        f"SELECT * FROM customer_mail_credentials WHERE account_id IN ({placeholders})", ids
    ).fetchall()
    return {int(row["account_id"]): dict(row) for row in rows}


def attach(account: Dict[str, Any], credential: Optional[Dict[str, Any]]) -> None:
    credential = credential or {}
    account["customer_secret"] = credential.get("secret") or ""
    account["customer_secret_enabled"] = credential.get("enabled", 0)
    account["customer_secret_expires_at"] = credential.get("expires_at")


def reset(account_id: int) -> Optional[str]:
    db = get_db()
    secret = generate_secret()
    try:
        cursor = db.execute(
            "UPDATE customer_mail_credentials SET secret = ? WHERE account_id = ?",
            (secret, int(account_id)),
        )
        if cursor.rowcount < 1:
            db.rollback()
            return None
        db.commit()
        return secret
    except Exception:
        db.rollback()
        return None


def update_access(account_id: int, *, enabled: bool, expires_at: Optional[str]) -> bool:
    db = get_db()
    cursor = db.execute(
        "UPDATE customer_mail_credentials SET enabled = ?, expires_at = ? WHERE account_id = ?",
        (int(enabled), expires_at, int(account_id)),
    )
    db.commit()
    return cursor.rowcount > 0
