from __future__ import annotations

import unittest
import os
import tempfile
from unittest.mock import patch

from tests._import_app import clear_login_attempts, import_web_app_module


class CustomerMailTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = import_web_app_module().app

    def setUp(self):
        with self.app.app_context():
            clear_login_attempts()
            from outlook_web.db import get_db

            db = get_db()
            db.execute("DELETE FROM accounts WHERE email LIKE '%@customer.test'")
            db.execute("DELETE FROM customer_mail_rate_limits")
            db.commit()
        self.client = self.app.test_client()

    def _create_account(self):
        from outlook_web.repositories import accounts

        with self.app.app_context():
            self.assertTrue(accounts.add_account("demo@customer.test", "password", "cid", "rt"))
            return accounts.get_account_by_email("demo@customer.test")

    def _query(self, secret, view="all"):
        return self.client.post(
            "/api/customer/messages",
            json={"email": "demo@customer.test", "secret": secret, "view": view},
        )

    def test_top_level_mailbox_compatibility_entry_is_removed(self):
        response = self.client.get("/mailbox")
        self.assertEqual(response.status_code, 404)

    def test_import_generates_secret_and_unknown_email_has_same_error(self):
        account = self._create_account()
        self.assertTrue(account["customer_secret"].startswith("mail_"))
        invalid = self._query("wrong")
        missing = self.client.post(
            "/api/customer/messages",
            json={"email": "missing@customer.test", "secret": "wrong", "view": "all"},
        )
        self.assertEqual(invalid.status_code, 403)
        self.assertEqual(invalid.get_json(), missing.get_json())

    def test_authorized_list_views_are_unfiltered_and_capped(self):
        account = self._create_account()
        messages = [{
            "id": str(index + 1),
            "subject": f"Message {index + 1}",
            "from_address": "sender@test.com",
            "created_at": "2026-10-01T00:00:00Z",
            "content_preview": f"Body {index + 1}",
        } for index in range(60)]
        with patch(
            "outlook_web.controllers.customer_mail.external_api.list_messages_for_customer",
            return_value=(messages, "Graph API"),
        ) as fetch:
            all_response = self._query(account["customer_secret"])
            junk_response = self._query(account["customer_secret"], "junk")
            invalid_response = self._query(account["customer_secret"], "verification")
        self.assertEqual(all_response.status_code, 200)
        self.assertEqual(len(all_response.get_json()["messages"]), 50)
        self.assertEqual(len(junk_response.get_json()["messages"]), 50)
        self.assertEqual(invalid_response.status_code, 400)
        self.assertEqual(fetch.call_args_list[0].kwargs["folder"], "inbox")
        self.assertEqual(fetch.call_args_list[0].kwargs["top"], 50)
        self.assertEqual(fetch.call_args.kwargs["folder"], "junkemail")
        self.assertEqual(fetch.call_args.kwargs["top"], 50)
        self.assertEqual(fetch.call_count, 2)
        self.assertNotIn("refresh_token", str(all_response.get_json()))
        self.assertNotIn("verification_code", str(all_response.get_json()))
        self.assertEqual(all_response.headers["Cache-Control"], "no-store")

    def test_missing_view_defaults_to_all_inbox_messages(self):
        account = self._create_account()
        with patch(
            "outlook_web.controllers.customer_mail.external_api.list_messages_for_customer",
            return_value=([], "Graph API"),
        ) as fetch:
            response = self.client.post(
                "/api/customer/messages",
                json={"email": account["email"], "secret": account["customer_secret"]},
            )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(fetch.call_args.kwargs["folder"], "inbox")
        self.assertEqual(fetch.call_args.kwargs["top"], 50)

    def test_regular_mail_and_same_code_messages_are_not_filtered_or_deduplicated(self):
        account = self._create_account()
        messages = [
            {
                "id": str(index + 1),
                "subject": subject,
                "from_address": "sender@test.com",
                "created_at": "2026-10-01T00:00:00Z",
                "content_preview": preview,
            }
            for index, (subject, preview) in enumerate([
                ("Weekly update", "Ordinary mail with no verification code"),
                ("Your verification code", "Code 123456"),
                ("Your verification code", "Code 123456"),
            ])
        ]
        with patch(
            "outlook_web.controllers.customer_mail.external_api.list_messages_for_customer",
            return_value=(messages, "Graph API"),
        ):
            response = self._query(account["customer_secret"])
        self.assertEqual(response.status_code, 200)
        result = response.get_json()["messages"]
        self.assertEqual(len(result), 3)
        self.assertEqual([item["subject"] for item in result], [item["subject"] for item in messages])
        self.assertEqual(len({item["message_id"] for item in result}), 3)
        self.assertTrue(all("verification_code" not in item for item in result))

    def test_reset_and_disable_revoke_access(self):
        account = self._create_account()
        with self.app.app_context():
            from outlook_web.repositories import accounts

            new_secret = accounts.reset_customer_secret(account["id"])
        self.assertNotEqual(new_secret, account["customer_secret"])
        self.assertEqual(self._query(account["customer_secret"]).status_code, 403)
        with self.app.app_context():
            accounts.update_customer_secret_access(account["id"], enabled=False, expires_at=None)
        self.assertEqual(self._query(new_secret).status_code, 403)
        with self.app.app_context():
            newest_secret = accounts.reset_customer_secret(account["id"])
        self.assertEqual(self._query(newest_secret).status_code, 403)

    def test_expired_secret_is_rejected(self):
        account = self._create_account()
        with self.app.app_context():
            from outlook_web.repositories import accounts

            accounts.update_customer_secret_access(
                account["id"], enabled=True, expires_at="2020-01-01T00:00:00Z"
            )
        self.assertEqual(self._query(account["customer_secret"]).status_code, 403)
        with self.app.app_context():
            rotated = accounts.reset_customer_secret(account["id"])
        self.assertEqual(self._query(rotated).status_code, 403)

    def test_oversized_payload_rejected_before_parsing(self):
        response = self.client.post(
            "/api/customer/messages",
            data='{"email":"' + "A" * 10000 + '"}',
            content_type="application/json",
        )
        self.assertEqual(response.status_code, 413)

    def test_case_sensitive_duplicate_addresses_have_distinct_secrets(self):
        self._create_account()
        from outlook_web.repositories import accounts

        with self.app.app_context():
            self.assertTrue(accounts.add_account("DEMO@customer.test", "", "cid", "rt"))
            second = accounts.get_account_by_email("DEMO@customer.test")
        with patch(
            "outlook_web.controllers.customer_mail.external_api.list_messages_for_customer",
            return_value=([], "Graph API"),
        ) as fetch:
            response = self.client.post(
                "/api/customer/messages",
                json={"email": "DEMO@customer.test", "secret": second["customer_secret"], "view": "all"},
            )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(fetch.call_args.kwargs["account"]["id"], second["id"])

    def test_unexpected_upstream_error_does_not_leak_credentials(self):
        account = self._create_account()
        with (
            patch(
                "outlook_web.controllers.customer_mail.external_api.list_messages_for_customer",
                side_effect=RuntimeError("private-token"),
            ),
            self.assertLogs("outlook_web.controllers.customer_mail", level="ERROR") as captured,
        ):
            response = self._query(account["customer_secret"])
        self.assertEqual(response.status_code, 502)
        self.assertNotIn("private-token", response.get_data(as_text=True))
        self.assertNotIn("private-token", "\n".join(captured.output))

    def test_graph_read_persists_rotated_refresh_token(self):
        account = self._create_account()
        with patch(
            "outlook_web.services.external_api.graph_service.get_emails_graph",
            return_value={
                "success": True,
                "new_refresh_token": "rotated-rt",
                "emails": [{"subject": "Graph message", "bodyPreview": "hello"}],
            },
        ):
            response = self._query(account["customer_secret"])
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["messages"][0]["subject"], "Graph message")
        with self.app.app_context():
            from outlook_web.repositories import accounts

            self.assertEqual(accounts.get_account_by_email("demo@customer.test")["refresh_token"], "rotated-rt")

    def test_graph_failure_uses_imap_fallback(self):
        account = self._create_account()
        with (
            patch(
                "outlook_web.services.external_api.graph_service.get_emails_graph",
                return_value={"success": False, "error": "unavailable"},
            ),
            patch(
                "outlook_web.services.external_api.imap_service.get_emails_imap_with_server",
                return_value={"success": True, "emails": [{"subject": "IMAP message", "content_preview": "hello"}]},
            ) as imap,
        ):
            response = self._query(account["customer_secret"], "junk")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["messages"][0]["subject"], "IMAP message")
        self.assertEqual(imap.call_args.args[3], "junkemail")

    def test_admin_secret_endpoints_require_login(self):
        account = self._create_account()
        self.assertEqual(
            self.client.post(f"/api/accounts/{account['id']}/customer-secret/reset").status_code,
            401,
        )
        self.assertEqual(
            self.client.patch(
                f"/api/accounts/{account['id']}/customer-secret",
                json={"enabled": False},
            ).status_code,
            401,
        )

    def test_admin_can_read_reset_and_disable_secret_without_caching(self):
        account = self._create_account()
        login = self.client.post("/login", json={"password": "testpass123"})
        self.assertEqual(login.status_code, 200)
        detail = self.client.get(f"/api/accounts/{account['id']}")
        listing = self.client.get("/api/accounts")
        self.assertEqual(detail.get_json()["account"]["customer_secret"], account["customer_secret"])
        self.assertEqual(detail.headers["Cache-Control"], "no-store")
        self.assertEqual(listing.headers["Cache-Control"], "no-store")
        reset = self.client.post(f"/api/accounts/{account['id']}/customer-secret/reset")
        self.assertEqual(reset.status_code, 200)
        self.assertEqual(reset.headers["Cache-Control"], "no-store")
        self.assertNotEqual(reset.get_json()["customer_secret"], account["customer_secret"])
        disabled = self.client.patch(
            f"/api/accounts/{account['id']}/customer-secret",
            json={"enabled": False, "expires_at": None},
        )
        self.assertEqual(disabled.status_code, 200)
        self.assertEqual(self._query(reset.get_json()["customer_secret"]).status_code, 403)

    @patch("outlook_web.controllers.customer_mail.time.time", return_value=1_800_000_000)
    def test_rate_limit_counts_failed_attempts_and_spoofed_ip_is_ignored(self, _clock):
        self._create_account()
        for index in range(30):
            response = self.client.post(
                "/api/customer/messages",
                json={"email": "demo@customer.test", "secret": "wrong"},
                headers={"X-Forwarded-For": f"192.0.2.{index + 1}"},
            )
            self.assertEqual(response.status_code, 403)
        response = self.client.post(
            "/api/customer/messages",
            json={"email": "demo@customer.test", "secret": "wrong"},
            headers={"X-Forwarded-For": "192.0.2.99"},
        )
        self.assertEqual(response.status_code, 429)
        self.assertEqual(response.headers["Cache-Control"], "no-store")

    def test_trusted_proxy_uses_last_forwarded_hop(self):
        from outlook_web.security.auth import get_client_ip

        with patch.dict(os.environ, {"TRUSTED_PROXIES": "127.0.0.1"}):
            with self.app.test_request_context(
                "/mailbox",
                headers={"X-Forwarded-For": "192.0.2.3, 198.51.100.4"},
                environ_base={"REMOTE_ADDR": "127.0.0.1"},
            ):
                self.assertEqual(get_client_ip(), "198.51.100.4")
            with self.app.test_request_context(
                "/mailbox",
                headers={"X-Forwarded-For": "192.0.2.3"},
                environ_base={"REMOTE_ADDR": "192.0.2.99"},
                environ_overrides={
                    "werkzeug.proxy_fix.orig": {"REMOTE_ADDR": "192.0.2.99"},
                    "REMOTE_ADDR": "192.0.2.3",
                },
            ):
                self.assertEqual(get_client_ip(), "192.0.2.99")

    def test_backfill_is_stable(self):
        account = self._create_account()
        with self.app.app_context():
            from outlook_web.db import get_db
            from outlook_web.repositories.customer_credentials import backfill_credentials

            db = get_db()
            db.execute("DELETE FROM customer_mail_credentials WHERE account_id = ?", (account["id"],))
            backfill_credentials(db)
            first = db.execute(
                "SELECT secret FROM customer_mail_credentials WHERE account_id = ?", (account["id"],)
            ).fetchone()[0]
            backfill_credentials(db)
            second = db.execute(
                "SELECT secret FROM customer_mail_credentials WHERE account_id = ?", (account["id"],)
            ).fetchone()[0]
            self.assertEqual(first, second)
            self.assertTrue(first.startswith("mail_"))

    def test_schema_has_no_customer_columns_and_delete_cascades(self):
        account = self._create_account()
        with self.app.app_context():
            from outlook_web.db import get_db

            db = get_db()
            columns = {row["name"] for row in db.execute("PRAGMA table_info(accounts)")}
            self.assertNotIn("customer_secret", columns)
            rows = db.execute(
                "SELECT secret FROM customer_mail_credentials WHERE account_id = ?", (account["id"],)
            ).fetchall()
            self.assertEqual(len(rows), 1)
            db.execute("DELETE FROM accounts WHERE id = ?", (account["id"],))
            self.assertIsNone(
                db.execute(
                    "SELECT secret FROM customer_mail_credentials WHERE account_id = ?", (account["id"],)
                ).fetchone()
            )
            db.commit()

    def test_batch_import_rollback_removes_credential(self):
        with self.app.app_context():
            from outlook_web.db import get_db
            from outlook_web.repositories import accounts

            db = get_db()
            self.assertTrue(accounts.add_account("rollback@customer.test", "", "cid", "rt", db=db, commit=False))
            self.assertEqual(
                db.execute(
                    "SELECT COUNT(*) FROM customer_mail_credentials c "
                    "JOIN accounts a ON a.id = c.account_id WHERE a.email = 'rollback@customer.test'"
                ).fetchone()[0],
                1,
            )
            db.rollback()
            self.assertIsNone(db.execute("SELECT id FROM accounts WHERE email = 'rollback@customer.test'").fetchone())

    def test_provisional_accounts_columns_migrate_without_rotating_secret(self):
        from outlook_web.db import init_db, create_sqlite_connection

        with tempfile.TemporaryDirectory() as temp:
            path = os.path.join(temp, "legacy.db")
            init_db(path)
            db = create_sqlite_connection(path)
            db.execute("ALTER TABLE accounts ADD COLUMN customer_secret TEXT")
            db.execute("ALTER TABLE accounts ADD COLUMN customer_secret_enabled INTEGER NOT NULL DEFAULT 1")
            db.execute("ALTER TABLE accounts ADD COLUMN customer_secret_expires_at TIMESTAMP")
            db.execute(
                "INSERT INTO accounts (email, client_id, refresh_token, customer_secret, "
                "customer_secret_enabled, customer_secret_expires_at) VALUES (?, ?, ?, ?, ?, ?)",
                ("legacy@customer.test", "cid", "rt", "mail_legacy", 0, "2030-01-01"),
            )
            db.execute("UPDATE settings SET value = '26' WHERE key = 'db_schema_version'")
            db.execute("DELETE FROM settings WHERE key = 'customer_mail_schema_version'")
            db.commit()
            db.close()
            init_db(path)
            db = create_sqlite_connection(path)
            row = db.execute(
                "SELECT c.* FROM customer_mail_credentials c "
                "JOIN accounts a ON a.id = c.account_id WHERE a.email = ?",
                ("legacy@customer.test",),
            ).fetchone()
            self.assertEqual((row["secret"], row["enabled"], row["expires_at"]), ("mail_legacy", 0, "2030-01-01"))
            self.assertEqual(db.execute("SELECT value FROM settings WHERE key = 'db_schema_version'").fetchone()[0], "25")
            init_db(path)
            self.assertEqual(
                db.execute("SELECT secret FROM customer_mail_credentials WHERE account_id = ?", (row["account_id"],))
                .fetchone()[0],
                "mail_legacy",
            )
            db.close()

    def test_legacy_column_cannot_resurrect_rotated_secret(self):
        from outlook_web.db import init_db, create_sqlite_connection

        with tempfile.TemporaryDirectory() as temp:
            path = os.path.join(temp, "legacy.db")
            init_db(path)
            db = create_sqlite_connection(path)
            db.execute("ALTER TABLE accounts ADD COLUMN customer_secret TEXT")
            db.execute(
                "INSERT INTO accounts (email, client_id, refresh_token, customer_secret) "
                "VALUES ('old@customer.test', 'cid', 'rt', 'mail_revoked')"
            )
            db.execute("DELETE FROM settings WHERE key = 'customer_mail_schema_version'")
            db.commit()
            db.close()
            init_db(path)
            db = create_sqlite_connection(path)
            account_id = db.execute("SELECT id FROM accounts WHERE email = 'old@customer.test'").fetchone()[0]
            self.assertEqual(
                db.execute("SELECT secret FROM customer_mail_credentials WHERE account_id = ?", (account_id,))
                .fetchone()[0],
                "mail_revoked",
            )
            db.execute(
                "UPDATE customer_mail_credentials SET secret = 'mail_rotated' WHERE account_id = ?",
                (account_id,),
            )
            db.execute("DELETE FROM customer_mail_credentials WHERE account_id = ?", (account_id,))
            db.commit()
            db.close()
            init_db(path)
            db = create_sqlite_connection(path)
            restored = db.execute(
                "SELECT secret FROM customer_mail_credentials WHERE account_id = ?", (account_id,)
            ).fetchone()[0]
            self.assertNotEqual(restored, "mail_revoked")
            self.assertTrue(restored.startswith("mail_"))
            db.close()


if __name__ == "__main__":
    unittest.main()
