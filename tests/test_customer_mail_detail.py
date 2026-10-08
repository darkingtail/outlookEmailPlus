from __future__ import annotations

import unittest
from unittest.mock import Mock, patch

from tests import test_customer_mail as base_tests


class CustomerDetailTests(unittest.TestCase):
    setUpClass = classmethod(base_tests.CustomerMailTests.setUpClass.__func__)
    setUp = base_tests.CustomerMailTests.setUp
    _create_account = base_tests.CustomerMailTests._create_account
    _query = base_tests.CustomerMailTests._query

    def _reference(self, account, *, method="Graph API", message_id="graph-message-1", view="all"):
        summary = {
            "id": message_id, "subject": "Your verification code", "from_address": "sender@example.test",
            "created_at": "2026-10-08T00:00:00Z", "content_preview": "Short preview",
        }
        with patch("outlook_web.services.external_api.list_messages_for_customer", return_value=([summary], method)):
            response = self._query(account["customer_secret"], view)
        self.assertEqual(response.status_code, 200)
        item = response.get_json()["messages"][0]
        return item.get("message_id", message_id)

    def test_list_emits_customer_detail_reference_and_detail_returns_full_body(self):
        account = self._create_account()
        reference = self._reference(account)
        self.assertNotEqual(reference, "graph-message-1")
        upstream = self._graph()
        with (
            patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result",
                  return_value={"success": True, "access_token": "fixture-at"}),
            patch("outlook_web.services.external_api.graph_service.requests.get", return_value=upstream),
        ):
            response = self._detail(account, reference)
        self.assertEqual(response.status_code, 200)
        self.assertGreater(len(response.get_json()["message"]["body"]), 1000)

    def _detail(self, account, reference, **overrides):
        payload = {"email": account["email"], "secret": account["customer_secret"], "message_id": reference}
        payload.update(overrides)
        return self.client.post("/api/customer/messages/detail", json=payload)

    def _graph(self, body=None, status=200):
        response = Mock(status_code=status)
        response.json.return_value = {
            "id": "graph-message-1", "subject": "Your verification code",
            "from": {"emailAddress": {"name": "Sender", "address": "sender@example.test"}},
            "toRecipients": [{"emailAddress": {"address": "demo@customer.test"}}],
            "ccRecipients": [{"emailAddress": {"address": "cc@example.test"}}],
            "receivedDateTime": "2026-10-08T00:00:00Z",
            "body": body or {"contentType": "text", "content": "Full paragraph\nCode 123456\n" + "tail " * 500},
            "bodyPreview": "Short preview",
            "raw_content": "PRIVATE-MIME", "attachments": ["PRIVATE-ATTACHMENT"],
            "refresh_token": "PRIVATE-RT", "bccRecipients": ["PRIVATE-BCC"],
        }
        return response

    def test_complete_graph_detail_is_allowlisted_and_never_fetches_mime(self):
        account = self._create_account()
        reference = self._reference(account)
        upstream = self._graph()
        with (
            patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result",
                  return_value={"success": True, "access_token": "fixture-at", "new_refresh_token": "rotated-detail-rt"}),
            patch("outlook_web.services.external_api.graph_service.requests.get", return_value=upstream) as read,
            patch("outlook_web.services.external_api.graph_service.get_email_raw_graph") as raw,
        ):
            response = self._detail(account, reference)
        self.assertEqual(response.status_code, 200)
        message = response.get_json()["message"]
        self.assertEqual(message["body"], upstream.json.return_value["body"]["content"])
        self.assertEqual(message["to"], "demo@customer.test")
        self.assertEqual(message["cc"], "cc@example.test")
        self.assertEqual(set(message), {"subject", "from", "to", "cc", "date", "body"})
        self.assertNotIn("PRIVATE", response.get_data(as_text=True))
        self.assertEqual(response.headers["Cache-Control"], "no-store")
        self.assertEqual(response.headers["Referrer-Policy"], "no-referrer")
        self.assertIn("/me/mailFolders/inbox/messages/graph-message-1", read.call_args.args[0])
        raw.assert_not_called()
        from outlook_web.repositories import accounts
        with self.app.app_context():
            self.assertEqual(accounts.get_account_by_email(account["email"])["refresh_token"], "rotated-detail-rt")

    def test_html_to_plaintext_preserves_full_content_without_active_elements(self):
        account = self._create_account()
        reference = self._reference(account)
        html = "<head><style>SECRET-CSS</style></head><p>First &amp; second</p><script>SECRET-SCRIPT</script>"
        html += '<p>Final paragraph <a href="https://example.test/verify">Verify</a></p><img src="https://tracker.test">'
        with (
            patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result",
                  return_value={"success": True, "access_token": "fixture-at"}),
            patch("outlook_web.services.external_api.graph_service.requests.get",
                  return_value=self._graph({"contentType": "html", "content": html})) as read,
        ):
            response = self._detail(account, reference)
        body = response.get_json()["message"]["body"]
        self.assertIn("First & second", body)
        self.assertIn("Final paragraph", body)
        self.assertIn("\n", body)
        self.assertNotIn("<", body)
        self.assertNotIn("SECRET", body)
        self.assertNotIn("tracker.test", body)
        self.assertEqual(read.call_count, 1)

    def test_graph_requests_html_to_avoid_expanded_tracking_links(self):
        account = self._create_account()
        reference = self._reference(account)
        tracking_url = "https://tracker.example.test/ls/click?opaque=" + "x" * 800
        html = (
            f'<p><a href="{tracking_url}">Mailbox</a></p>'
            "<p>Enter this code to continue: 918273</p>"
            "<p>If you did not request this, ignore this email.</p>"
            f'<p><a href="{tracking_url}">Help center</a></p>'
            "<p>Final paragraph</p>"
        )
        converted_text = f"Mailbox<{tracking_url}>\nEnter this code: 918273\nHelp center<{tracking_url}>"

        def read_graph(*args, **kwargs):
            if kwargs["headers"].get("Prefer") == 'outlook.body-content-type="html"':
                return self._graph({"contentType": "html", "content": html})
            return self._graph({"contentType": "text", "content": converted_text})

        with (
            patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result",
                  return_value={"success": True, "access_token": "fixture-at"}),
            patch("outlook_web.services.external_api.graph_service.requests.get", side_effect=read_graph) as read,
        ):
            response = self._detail(account, reference)
        self.assertEqual(response.status_code, 200)
        body = response.get_json()["message"]["body"]
        self.assertNotIn("tracker.example.test", body)
        self.assertEqual(
            body,
            "Mailbox\n\nEnter this code to continue: 918273\n\n"
            "If you did not request this, ignore this email.\n\nHelp center\n\nFinal paragraph",
        )
        self.assertEqual(read.call_args.kwargs["headers"]["Prefer"], 'outlook.body-content-type="html"')

    def test_wrong_key_unknown_email_and_other_account_cannot_read(self):
        account = self._create_account()
        reference = self._reference(account)
        from outlook_web.repositories import accounts
        with self.app.app_context():
            accounts.add_account("second@customer.test", "", "cid", "rt")
            other = accounts.get_account_by_email("second@customer.test")
        with patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result") as token:
            wrong = self._detail(account, reference, secret="wrong")
            unknown = self._detail(account, reference, email="missing@customer.test", secret="wrong")
            crossed = self._detail(other, reference)
        self.assertEqual(wrong.status_code, 403)
        self.assertEqual(wrong.get_json(), unknown.get_json())
        self.assertEqual(crossed.status_code, 400)
        token.assert_not_called()

    def test_revoked_expired_disabled_and_inactive_account_are_checked_on_detail(self):
        account = self._create_account()
        reference = self._reference(account)
        from outlook_web.repositories import accounts
        from outlook_web.db import get_db
        with patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result") as token:
            with self.app.app_context():
                accounts.update_customer_secret_access(account["id"], enabled=False, expires_at=None)
            self.assertEqual(self._detail(account, reference).status_code, 403)
            with self.app.app_context():
                accounts.update_customer_secret_access(account["id"], enabled=True, expires_at="2020-01-01")
            self.assertEqual(self._detail(account, reference).status_code, 403)
            with self.app.app_context():
                accounts.update_customer_secret_access(account["id"], enabled=True, expires_at=None)
                get_db().execute("UPDATE accounts SET status = 'inactive' WHERE id = ?", (account["id"],))
                get_db().commit()
            self.assertEqual(self._detail(account, reference).status_code, 403)
            with self.app.app_context():
                get_db().execute("UPDATE accounts SET status = 'active' WHERE id = ?", (account["id"],))
                get_db().commit()
                accounts.reset_customer_secret(account["id"])
            self.assertEqual(self._detail(account, reference).status_code, 403)
            token.assert_not_called()

    def test_tampered_expired_and_unsigned_message_ids_are_rejected(self):
        account = self._create_account()
        reference = self._reference(account)
        with patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result") as token:
            for value in ("graph-message-1", reference + "x", "../other", "*", "1:*", "%2f"):
                with self.subTest(value=value[:10]):
                    self.assertEqual(self._detail(account, value).status_code, 400)
            with patch("itsdangerous.timed.TimestampSigner.get_timestamp", return_value=1):
                expired = self._reference(account)
            self.assertEqual(self._detail(account, expired).status_code, 400)
            token.assert_not_called()

    def test_folder_and_method_are_bound_to_signed_reference(self):
        account = self._create_account()
        reference = self._reference(account, view="junk")
        with (
            patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result",
                  return_value={"success": True, "access_token": "fixture-at"}),
            patch("outlook_web.services.external_api.graph_service.requests.get", return_value=self._graph()) as read,
        ):
            response = self._detail(account, reference, folder="deleteditems", method="IMAP (Old)")
        self.assertEqual(response.status_code, 200)
        self.assertIn("/mailFolders/junkemail/messages/", read.call_args.args[0])

    def test_upstream_failures_and_missing_mail_are_sanitized(self):
        account = self._create_account()
        reference = self._reference(account)
        for upstream_status, expected in ((404, 404), (401, 502), (500, 502)):
            with (
                self.subTest(status=upstream_status),
                patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result",
                      return_value={"success": True, "access_token": "fixture-at"}),
                patch("outlook_web.services.external_api.graph_service.requests.get",
                      return_value=self._graph(status=upstream_status)),
            ):
                response = self._detail(account, reference)
                self.assertEqual(response.status_code, expected)
                self.assertEqual(response.headers["Cache-Control"], "no-store")
                self.assertNotIn("PRIVATE", response.get_data(as_text=True))
        with (
            patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result",
                  side_effect=RuntimeError("PRIVATE-TOKEN")),
            self.assertLogs("outlook_web.controllers.customer_mail", level="ERROR") as logs,
        ):
            response = self._detail(account, reference)
        self.assertEqual(response.status_code, 502)
        self.assertNotIn("PRIVATE-TOKEN", "\n".join(logs.output) + response.get_data(as_text=True))

    def test_outlook_imap_detail_does_not_switch_protocol_and_detects_sequence_drift(self):
        account = self._create_account()
        reference = self._reference(account, method="IMAP (Old)", message_id="12", view="junk")
        detail = {
            "id": "12", "subject": "Your verification code", "from": "sender@example.test",
            "date": "2026-10-08T00:00:00Z", "to": "demo@customer.test", "cc": "cc@example.test",
            "body": "<p>Complete IMAP body</p>", "body_type": "html", "raw_content": "PRIVATE-MIME",
        }
        with (
            patch("outlook_web.services.external_api.imap_service.get_email_detail_imap_with_server",
                  return_value=detail) as read,
            patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result") as graph,
        ):
            response = self._detail(account, reference)
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.get_json()["message"]["body"], "Complete IMAP body")
            self.assertEqual(read.call_args.args[3:], ("12", "junkemail", "outlook.office365.com"))
            detail["subject"] = "Different message after deletion"
            self.assertEqual(self._detail(account, reference).status_code, 404)
            graph.assert_not_called()

    def test_outlook_imap_customer_prefers_html_without_changing_admin_default(self):
        from email.message import EmailMessage
        from outlook_web.services import imap

        account = self._create_account()
        reference = self._reference(account, method="IMAP (New)", message_id="12")
        tracking_url = "https://tracker.example.test/ls/click?opaque=" + "x" * 800
        text = f"Enter this code: 918273\nHelp center<{tracking_url}>\nFinal paragraph"
        message = EmailMessage()
        message["Subject"] = "Your verification code"
        message["From"] = "sender@example.test"
        message["To"] = "demo@customer.test"
        message["Date"] = "Thu, 08 Oct 2026 00:00:00 +0000"
        message.set_content(text)
        message.add_alternative(
            f'<p>Enter this code: 918273</p><p><a href="{tracking_url}">Help center</a></p>'
            "<p>Final paragraph</p>",
            subtype="html",
        )
        message.add_attachment(b"<p>PRIVATE-ATTACHMENT</p>", maintype="text", subtype="html", filename="attachment.html")
        connection = Mock()
        connection.select.return_value = ("OK", [])
        connection.fetch.return_value = ("OK", [(b"12 (RFC822)", message.as_bytes())])
        with (
            patch("outlook_web.services.imap.get_access_token_imap", return_value="fixture-at"),
            patch("outlook_web.services.imap.imaplib.IMAP4_SSL", return_value=connection),
        ):
            response = self._detail(account, reference)
            admin_detail = imap.get_email_detail_imap_with_server(
                account["email"], "fixture-cid", "fixture-rt", "12",
            )
        self.assertEqual(response.status_code, 200)
        body = response.get_json()["message"]["body"]
        self.assertEqual(body, "Enter this code: 918273\n\nHelp center\n\nFinal paragraph")
        self.assertNotIn("PRIVATE", body)
        self.assertNotIn("tracker.example.test", body)
        self.assertEqual(admin_detail["body_type"], "text")
        self.assertIn(tracking_url, admin_detail["body"].replace("\n", ""))

    def test_generic_imap_uses_bound_account_and_uid(self):
        account = self._create_account()
        from outlook_web.db import get_db
        with self.app.app_context():
            get_db().execute("UPDATE accounts SET account_type='imap', imap_host='fixture.test', "
                             "imap_password='fixture-pass' WHERE id=?", (account["id"],))
            get_db().commit()
        reference = self._reference(account, method="IMAP (Generic)", message_id="55")
        with patch("outlook_web.services.external_api.get_email_detail_imap_generic_result", return_value={
            "success": True, "email": {"id": "55", "subject": "Your verification code",
                                      "from": "sender@example.test", "date": "2026-10-08T00:00:00Z",
                                      "body_text": "Plaintext alternative", "body_html": "<p>Complete generic body</p>",
                                      "raw_content": "PRIVATE-MIME", "to": "demo@customer.test", "cc": "cc@test"},
        }) as read:
            response = self._detail(account, reference, account_id=987, email_addr="someone@customer.test")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["message"]["body"], "Complete generic body")
        self.assertEqual(read.call_args.kwargs["message_id"], "55")
        self.assertEqual(read.call_args.kwargs["email_addr"], "demo@customer.test")

    def test_outlook_imap_empty_html_keeps_plaintext_fallback(self):
        from email.message import EmailMessage
        from outlook_web.services import imap
        from outlook_web.services.customer_mail import _plain_body

        message = EmailMessage()
        text = "Complete plaintext\nVisit https://example.test/help\nFinal paragraph"
        message.set_content(text)
        message.add_alternative(
            "<html><head><style>p { color: red; }</style></head><body></body></html>",
            subtype="html",
        )
        connection = Mock()
        connection.select.return_value = ("OK", [])
        connection.fetch.return_value = ("OK", [(b"12 (RFC822)", message.as_bytes())])
        with (
            patch("outlook_web.services.imap.get_access_token_imap", return_value="fixture-at"),
            patch("outlook_web.services.imap.imaplib.IMAP4_SSL", return_value=connection),
        ):
            detail = imap.get_email_detail_imap_with_server(
                "demo@customer.test", "fixture-cid", "fixture-rt", "12", prefer_html=True,
            )
        self.assertEqual(_plain_body(detail).strip(), text)

    @patch("outlook_web.controllers.customer_mail.time.time", return_value=1_800_000_000)
    def test_detail_limits_input_and_shares_list_rate_limit(self, _clock):
        account = self._create_account()
        reference = self._reference(account)
        for secret in ("密钥", [], {"key": "wrong"}):
            response = self._detail(account, reference, secret=secret)
            self.assertIn(response.status_code, (400, 403))
        oversize = self.client.post("/api/customer/messages/detail", json={"email": "x" * 10000})
        self.assertEqual(oversize.status_code, 413)
        for _ in range(30):
            response = self._detail(account, reference, secret="wrong")
        self.assertEqual(response.status_code, 429)
        self.assertEqual(response.headers["Cache-Control"], "no-store")
        self.assertIn(self.client.get("/api/customer/messages/detail").status_code, (405, 401))

    def test_reset_invalidates_old_reference_even_with_new_credential(self):
        account = self._create_account()
        reference = self._reference(account)
        from outlook_web.repositories import accounts
        with self.app.app_context():
            account["customer_secret"] = accounts.reset_customer_secret(account["id"])
        with patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result") as token:
            self.assertEqual(self._detail(account, reference).status_code, 400)
        token.assert_not_called()

    def test_rotated_graph_token_is_saved_even_when_message_read_fails(self):
        account = self._create_account()
        reference = self._reference(account)
        with (
            patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result",
                  return_value={"success": True, "access_token": "fixture-at", "refresh_token": "rotated-on-error"}),
            patch("outlook_web.services.external_api.graph_service.requests.get", return_value=self._graph(status=500)),
        ):
            self.assertEqual(self._detail(account, reference).status_code, 502)
        from outlook_web.repositories import accounts
        with self.app.app_context():
            self.assertEqual(accounts.get_account_by_email(account["email"])["refresh_token"], "rotated-on-error")

    def test_graph_detail_failure_never_falls_back_to_imap(self):
        account = self._create_account()
        reference = self._reference(account)
        with (
            patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result",
                  return_value={"success": False, "error": "PRIVATE"}),
            patch("outlook_web.services.external_api.imap_service.get_email_detail_imap_with_server") as imap,
        ):
            response = self._detail(account, reference)
        self.assertEqual(response.status_code, 502)
        self.assertNotIn("PRIVATE", response.get_data(as_text=True))
        imap.assert_not_called()

    def test_mismatched_graph_message_identity_is_not_returned(self):
        account = self._create_account()
        reference = self._reference(account)
        for field, value in (("id", "another-id"), ("subject", "different"), ("receivedDateTime", "2020-01-01T00:00:00Z"),
                             ("from", {"emailAddress": {"address": "someone-else@test"}})):
            with self.subTest(field=field):
                upstream = self._graph()
                upstream.json.return_value[field] = value
                with (
                    patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result",
                          return_value={"success": True, "access_token": "fixture-at"}),
                    patch("outlook_web.services.external_api.graph_service.requests.get", return_value=upstream),
                ):
                    response = self._detail(account, reference)
                self.assertEqual(response.status_code, 404)
                self.assertNotIn("tail", response.get_data(as_text=True))

    def test_folder_channel_and_id_validation_rejects_unsafe_signed_values(self):
        account = self._create_account()
        from outlook_web.services import customer_mail
        self._reference(account)
        with self.app.app_context():
            serializer = customer_mail._serializer(account)
            valid = {"id": "123", "folder": "inbox", "method": "IMAP (New)", "identity": "a" * 64}
            bad_values = [dict(valid, id=x) for x in ("1:*", "1,2", "*", "0", "../x")]
            bad_values += [dict(valid, folder="deleteditems"), dict(valid, method="unknown"),
                           dict(valid, method="IMAP (Generic)"), dict(valid, identity=None)]
            bad_values += [dict(valid, method="Graph API", id="..")]
            references = [serializer.dumps(value) for value in bad_values]
        with patch("outlook_web.services.external_api.imap_service.get_email_detail_imap_with_server") as imap:
            for reference in references:
                self.assertEqual(self._detail(account, reference).status_code, 400)
        imap.assert_not_called()

    def test_graph_id_is_url_encoded_and_redirects_are_not_followed(self):
        account = self._create_account()
        reference = self._reference(account, message_id="a/b+==?foo=bar")
        upstream = self._graph()
        upstream.json.return_value["id"] = "a/b+==?foo=bar"
        with (
            patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result",
                  return_value={"success": True, "access_token": "fixture-at"}),
            patch("outlook_web.services.external_api.graph_service.requests.get", return_value=upstream) as read,
        ):
            self.assertEqual(self._detail(account, reference).status_code, 200)
        self.assertTrue(read.call_args.args[0].endswith("/a%2Fb%2B%3D%3D%3Ffoo%3Dbar"))
        self.assertFalse(read.call_args.kwargs["allow_redirects"])

    def test_html_hides_tracking_content_and_preserves_visible_link_text(self):
        from outlook_web.services.customer_mail import _plain_body
        body = _plain_body({"body_html": """
            <p>Visible text</p><div hidden>PRIVATE-HIDDEN</div>
            <span style="display: none !important">PRIVATE-TRACK</span>
            <script>PRIVATE-JS</script><img src="https://PRIVATE-TRACKER.test">
            <p><a href="https://verify.example.test/path">Verify account</a>
            <a href="javascript:PRIVATE">Bad link</a></p><p>Final</p>
        """})
        self.assertIn("Visible text", body)
        self.assertIn("Verify account", body)
        self.assertIn("Bad link", body)
        self.assertIn("Final", body)
        self.assertNotIn("PRIVATE", body)
        self.assertNotIn("<", body)

    def test_html_link_destinations_are_not_dumped_into_customer_body(self):
        from outlook_web.services.customer_mail import _plain_body

        body = _plain_body({
            "body_html": '<p>Verify your account <a href="https://verify.example.test/path?token=PRIVATE">here</a>.</p>'
        })
        self.assertIn("Verify your account here.", body)
        self.assertNotIn("verify.example.test", body)
        self.assertNotIn("PRIVATE", body)

    def test_empty_body_is_not_replaced_by_preview(self):
        account = self._create_account()
        reference = self._reference(account)
        with (
            patch("outlook_web.services.external_api.graph_service.get_access_token_graph_result",
                  return_value={"success": True, "access_token": "fixture-at"}),
            patch("outlook_web.services.external_api.graph_service.requests.get",
                  return_value=self._graph({"contentType": "text", "content": ""})),
        ):
            response = self._detail(account, reference)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["message"]["body"], "")

    def test_whitespace_plaintext_does_not_hide_complete_html_body(self):
        from outlook_web.services.customer_mail import _plain_body
        self.assertEqual(
            _plain_body({"body_text": "\r\n  \t", "body_html": "<p>Complete body</p>"}),
            "Complete body",
        )

    def test_unclosed_html_head_does_not_hide_body(self):
        from outlook_web.services.customer_mail import _plain_body
        self.assertEqual(
            _plain_body({"body_html": "<html><head><title>Notice</title><body><p>Complete body</p></body></html>"}),
            "Complete body",
        )

    def test_imap_html_alternative_is_preferred_to_expanded_plaintext_links(self):
        from outlook_web.services.customer_mail import _plain_body
        tracking_url = "https://tracker.example.test/ls/click?opaque=" + "x" * 800
        body = _plain_body({
            "body_text": f"Complete generic body\nHelp center<{tracking_url}>",
            "body_html": f'<p>Complete generic body</p><p><a href="{tracking_url}">Help center</a></p>',
        })
        self.assertEqual(body, "Complete generic body\n\nHelp center")
        self.assertNotIn("tracker.example.test", body)

    def test_empty_html_alternative_keeps_complete_plaintext(self):
        from outlook_web.services.customer_mail import _plain_body
        text = "Complete plaintext\nVisit https://example.test/help\nFinal paragraph"
        for html in ("", " \r\n\t", '<html><head><style>p { color: red; }</style></head><body></body></html>'):
            with self.subTest(html=html):
                self.assertEqual(_plain_body({"body_text": text, "body_html": html}), text)

    def test_real_plaintext_urls_are_preserved(self):
        from outlook_web.services.customer_mail import _plain_body
        text = "Visit https://example.test/help?case=123\nFinal paragraph"
        self.assertEqual(_plain_body({"body": {"contentType": "text", "content": text}}), text)
        self.assertEqual(_plain_body({"body": text, "body_type": "text"}), text)


if __name__ == "__main__":
    unittest.main()
