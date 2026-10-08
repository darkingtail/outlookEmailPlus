from __future__ import annotations

import hmac
import logging
import time
from datetime import datetime, timezone

from flask import jsonify, render_template, request

from outlook_web.audit import log_audit
from outlook_web.db import get_db
from outlook_web.repositories import accounts as accounts_repo
from outlook_web.security.auth import get_client_ip
from outlook_web.services import external_api
from outlook_web.services import customer_mail as customer_mail_service

_LOGGER = logging.getLogger(__name__)
_RATE_LIMIT_PER_MINUTE = 30
_MAX_INPUT_LENGTH = 254


def page():
    return render_template("customer_mail.html")


def _rate_limited() -> bool:
    db = get_db()
    bucket = int(time.time()) // 60
    ip = get_client_ip()
    db.execute(
        """
        INSERT INTO customer_mail_rate_limits (ip_address, minute_bucket, request_count)
        VALUES (?, ?, 1)
        ON CONFLICT(ip_address, minute_bucket)
        DO UPDATE SET request_count = request_count + 1
        """,
        (ip, bucket),
    )
    db.execute("DELETE FROM customer_mail_rate_limits WHERE minute_bucket < ?", (bucket - 5,))
    db.commit()
    row = db.execute(
        "SELECT request_count FROM customer_mail_rate_limits WHERE ip_address = ? AND minute_bucket = ?",
        (ip, bucket),
    ).fetchone()
    return bool(row and row["request_count"] > _RATE_LIMIT_PER_MINUTE)


def _authorized_account(email: str, secret: str):
    account = accounts_repo.get_customer_account(email)
    stored = str(account.get("customer_secret") or "") if account else ""
    if not account or not stored or not hmac.compare_digest(stored.encode("utf-8"), secret.encode("utf-8")):
        return None
    if not account.get("customer_secret_enabled") or account.get("status") != "active":
        return None
    expires_at = str(account.get("customer_secret_expires_at") or "").strip()
    if expires_at:
        try:
            expiration = datetime.fromisoformat(expires_at.replace("Z", "+00:00"))
            if expiration.tzinfo is None:
                expiration = expiration.replace(tzinfo=timezone.utc)
            if expiration <= datetime.now(timezone.utc):
                return None
        except ValueError:
            return None
    if account.get("_credential_errors"):
        return None
    return account


def _customer_request():
    # No credentials in URLs, response caches or exception logs.
    request.max_content_length = 4096
    try:
        if _rate_limited():
            return None, None, (jsonify({"success": False, "message": "请求过于频繁，请稍后重试"}), 429)
    except Exception:
        _LOGGER.exception("customer mail rate limit unavailable")
        return None, None, (jsonify({"success": False, "message": "服务暂不可用"}), 503)

    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return None, None, (jsonify({"success": False, "message": "请求参数无效"}), 400)
    email = payload.get("email")
    secret = payload.get("secret")
    if (
        not isinstance(email, str)
        or not isinstance(secret, str)
        or not email.strip()
        or "@" not in email
        or len(email) > _MAX_INPUT_LENGTH
        or not secret
        or len(secret) > 256
    ):
        return None, None, (jsonify({"success": False, "message": "请求参数无效"}), 400)
    email = email.strip()
    account = _authorized_account(email, secret)
    if account is None:
        log_audit("customer_mail_denied", "account", details="invalid credentials or access")
        return None, None, (jsonify({"success": False, "message": "邮箱或密钥无效"}), 403)
    return payload, account, None


def messages():
    payload, account, error = _customer_request()
    if error is not None:
        return error
    view = str(payload.get("view") or "all").strip().lower()
    if view not in {"all", "junk"}:
        return jsonify({"success": False, "message": "请求参数无效"}), 400

    folder = "junkemail" if view == "junk" else "inbox"
    try:
        messages_list, _method = external_api.list_messages_for_customer(
            account=account, folder=folder, top=50
        )
    except external_api.ExternalApiError:
        log_audit("customer_mail_failed", "account", str(account["id"]), "upstream read failed")
        return jsonify({"success": False, "message": "暂时无法读取邮件"}), 502
    except Exception as exc:
        _LOGGER.error("customer mail read failed: %s", type(exc).__name__)
        return jsonify({"success": False, "message": "暂时无法读取邮件"}), 502

    output = []
    for item in messages_list[:50]:
        subject = str(item.get("subject") or "")
        preview = str(item.get("content_preview") or "")
        output.append(
            {
                "message_id": customer_mail_service.issue_message_reference(account, item, folder=folder, method=_method),
                "subject": subject,
                "from": str(item.get("from_address") or ""),
                "date": str(item.get("created_at") or ""),
                "preview": preview,
            }
        )
    log_audit("customer_mail_read", "account", str(account["id"]), f"view={view}")
    response = jsonify({"success": True, "messages": output})
    response.headers["Cache-Control"] = "no-store"
    return response


def message_detail():
    payload, account, error = _customer_request()
    if error is not None:
        return error
    reference = payload.get("message_id")
    if not isinstance(reference, str) or not reference or len(reference) > 3072:
        return jsonify({"success": False, "message": "请求参数无效"}), 400
    try:
        message = customer_mail_service.get_message_detail(account, reference)
    except external_api.InvalidParamError:
        return jsonify({"success": False, "message": "邮件引用无效或已过期，请刷新邮件列表"}), 400
    except external_api.MailNotFoundError:
        return jsonify({"success": False, "message": "邮件已变化或不存在，请刷新邮件列表"}), 404
    except external_api.ExternalApiError:
        log_audit("customer_mail_detail_failed", "account", str(account["id"]), "upstream read failed")
        return jsonify({"success": False, "message": "暂时无法读取邮件详情，请稍后重试"}), 502
    except Exception as exc:
        _LOGGER.error("customer mail detail failed: %s", type(exc).__name__)
        return jsonify({"success": False, "message": "暂时无法读取邮件详情，请稍后重试"}), 502
    log_audit("customer_mail_detail_read", "account", str(account["id"]), "detail read")
    return jsonify({"success": True, "message": message})
