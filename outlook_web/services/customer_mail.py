"""Customer-only message references and safe, complete plaintext details."""
from __future__ import annotations

import hashlib
import json
import re
from html.parser import HTMLParser
from urllib.parse import quote

from flask import current_app
from itsdangerous import BadData, URLSafeTimedSerializer

from outlook_web.repositories import accounts as accounts_repo
from outlook_web.services import external_api

MESSAGE_REFERENCE_TTL = 15 * 60
_METHODS = {"Graph API", "IMAP (New)", "IMAP (Old)", "IMAP (Generic)"}
_FOLDERS = {"inbox", "junkemail"}


def _serializer(account):
    # Binding the current credential also invalidates old references after a key reset.
    binding = json.dumps(
        [account["id"], account["email"], account["customer_secret"]], ensure_ascii=True
    ).encode("utf-8")
    salt = "customer-mail-detail-v1:" + hashlib.sha256(binding).hexdigest()
    return URLSafeTimedSerializer(current_app.secret_key, salt=salt)


def _identity(summary):
    data = [summary.get("subject") or "无主题", summary.get("from_address") or "", summary.get("created_at") or ""]
    return hashlib.sha256(json.dumps(data, ensure_ascii=True).encode("utf-8")).hexdigest()


def _valid_id(message_id, method):
    if not isinstance(message_id, str) or not message_id or len(message_id) > 2048:
        return False
    if method == "Graph API":
        return message_id not in {".", ".."} and not any(ord(char) < 32 for char in message_id)
    return re.fullmatch(r"[1-9][0-9]{0,19}", message_id) is not None


def issue_message_reference(account, summary, *, folder, method):
    message_id = summary.get("id")
    if folder not in _FOLDERS or method not in _METHODS or not _valid_id(message_id, method):
        return ""
    return _serializer(account).dumps(
        {"id": message_id, "folder": folder, "method": method, "identity": _identity(summary)}
    )


def _resolve_reference(account, reference):
    try:
        value = _serializer(account).loads(reference, max_age=MESSAGE_REFERENCE_TTL)
    except BadData:
        raise external_api.InvalidParamError("邮件引用已失效，请刷新邮件列表") from None
    if (
        not isinstance(value, dict)
        or value.get("folder") not in _FOLDERS
        or value.get("method") not in _METHODS
        or not _valid_id(value.get("id"), value.get("method"))
        or not isinstance(value.get("identity"), str)
        or len(value["identity"]) != 64
    ):
        raise external_api.InvalidParamError("邮件引用无效，请刷新邮件列表")
    generic_account = str(account.get("account_type") or "outlook").lower() == "imap"
    if generic_account != (value["method"] == "IMAP (Generic)"):
        raise external_api.InvalidParamError("邮箱配置已变化，请刷新邮件列表")
    return value


class _PlainText(HTMLParser):
    """Never load HTML resources; preserve visible text and paragraph structure."""

    _SKIP = {"head", "script", "style", "template", "noscript", "svg", "math", "iframe", "object", "audio", "video", "canvas"}
    _VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"}
    _BLOCK = {"p", "div", "section", "article", "header", "footer", "li", "tr", "h1", "h2", "h3", "h4", "blockquote", "pre"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.blocked = []

    def handle_starttag(self, tag, attrs):
        # Email HTML may omit </head>; body content implicitly ends a bare head.
        # Do not release nested suppressed content such as scripts or templates.
        if self.blocked == ["head"] and (tag == "body" or tag in self._BLOCK):
            self.blocked.clear()
        attrs = dict(attrs)
        style = re.sub(r"\s+", "", attrs.get("style") or "").lower()
        hidden = "hidden" in attrs or (attrs.get("aria-hidden") or "").lower() == "true"
        hidden = hidden or "display:none" in style or "visibility:hidden" in style
        if self.blocked or tag in self._SKIP or hidden:
            if tag not in self._VOID:
                self.blocked.append(tag)
            return
        if tag in self._BLOCK or tag in {"br", "hr"}:
            self.parts.append("\n")
        elif tag in {"td", "th"}:
            self.parts.append("\t")

    def handle_endtag(self, tag):
        if self.blocked:
            if tag in self.blocked:
                index = len(self.blocked) - 1 - self.blocked[::-1].index(tag)
                del self.blocked[index:]
            return
        if tag in self._BLOCK:
            self.parts.append("\n")

    def handle_data(self, data):
        if not self.blocked:
            self.parts.append(data)

    def text(self):
        return re.sub(r"\n[ \t]*(?:\n[ \t]*)+", "\n\n", "".join(self.parts)).strip()


def _plain_body(detail):
    body = detail.get("body")
    if isinstance(body, dict):
        content = str(body.get("content") or "")
        is_html = str(body.get("contentType") or "").lower() == "html"
    elif str(detail.get("body_html") or "").strip():
        content = str(detail["body_html"])
        is_html = True
    elif str(detail.get("body_text") or "").strip():
        return str(detail["body_text"])
    else:
        content = str(detail.get("body_html") or body or "")
        is_html = bool(detail.get("body_html")) or str(detail.get("body_type") or "").lower() == "html"
    if not is_html:
        return content
    parser = _PlainText()
    parser.feed(content)
    parser.close()
    text = parser.text()
    if not text and not isinstance(body, dict) and str(detail.get("body_text") or "").strip():
        return str(detail["body_text"])
    return text


def _graph_detail(account, reference):
    graph = external_api.graph_service
    proxy_url = external_api._get_proxy_url(account)
    token = graph.get_access_token_graph_result(account.get("client_id") or "", account.get("refresh_token") or "", proxy_url)
    if not token.get("success") or not token.get("access_token"):
        raise external_api.UpstreamReadFailedError("暂时无法读取邮件详情")
    rotated = token.get("new_refresh_token") or token.get("refresh_token")
    if rotated:
        accounts_repo.update_refresh_token_if_changed(int(account["id"]), str(rotated))
    # Keep the folder restriction. Never switch protocols or fetch the raw MIME endpoint.
    message_id = quote(reference["id"], safe="")
    url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{reference['folder']}/messages/{message_id}"
    response = graph.requests.get(
        url,
        params={"$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body"},
        headers={"Authorization": f"Bearer {token['access_token']}", "Prefer": 'outlook.body-content-type="html"'},
        timeout=30,
        proxies=graph.build_proxies(proxy_url),
        allow_redirects=False,
    )
    if response.status_code == 404:
        raise external_api.MailNotFoundError("邮件已变化或不存在，请刷新邮件列表")
    if response.status_code != 200:
        raise external_api.UpstreamReadFailedError("暂时无法读取邮件详情")
    return response.json()


def get_message_detail(account, reference):
    reference = _resolve_reference(account, reference)
    method = reference["method"]
    if method == "Graph API":
        detail = _graph_detail(account, reference)
    elif method == "IMAP (Generic)":
        result = external_api.get_email_detail_imap_generic_result(
            email_addr=account["email"],
            imap_password=account.get("imap_password") or "",
            imap_host=account.get("imap_host") or "",
            imap_port=account.get("imap_port") or 993,
            provider=account.get("provider") or "_default",
            folder=reference["folder"],
            message_id=reference["id"],
        )
        if not result.get("success"):
            raise external_api.UpstreamReadFailedError("暂时无法读取邮件详情")
        detail = result.get("email")
    else:
        server = external_api.IMAP_SERVER_NEW if method == "IMAP (New)" else external_api.IMAP_SERVER_OLD
        detail = external_api.imap_service.get_email_detail_imap_with_server(
            account["email"], account.get("client_id") or "", account.get("refresh_token") or "",
            reference["id"], reference["folder"], server,
            prefer_html=True,
        )
        if not detail:
            raise external_api.UpstreamReadFailedError("暂时无法读取邮件详情")
    if not isinstance(detail, dict):
        raise external_api.UpstreamReadFailedError("暂时无法读取邮件详情")
    summary = external_api._build_message_summary(account["email"], detail, method=method)
    # Outlook OAuth IMAP currently uses sequence numbers: detect a different message
    # after expunge/move instead of returning a body under the previous subject.
    if str(detail.get("id") or "") != reference["id"] or _identity(summary) != reference["identity"]:
        raise external_api.MailNotFoundError("邮件已变化或不存在，请刷新邮件列表")
    body = _plain_body(detail)
    if method == "Graph API":
        def recipients(field):
            return ", ".join(str((item.get("emailAddress") or {}).get("address") or "") for item in detail.get(field) or [])
        to, cc = recipients("toRecipients"), recipients("ccRecipients")
    else:
        to, cc = str(detail.get("to") or ""), str(detail.get("cc") or "")
    # An explicit allowlist is required: upstream dictionaries may contain credentials,
    # raw MIME, attachments and HTML. None of those belong to the customer contract.
    return {
        "subject": summary["subject"], "from": summary["from_address"], "date": summary["created_at"],
        "to": to, "cc": cc, "body": body,
    }
