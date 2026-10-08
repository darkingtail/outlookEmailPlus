from __future__ import annotations

from typing import Callable, Optional

from flask import Blueprint

from outlook_web.controllers import pages as pages_controller


def create_blueprint(csrf_exempt: Optional[Callable] = None) -> Blueprint:
    """创建 pages Blueprint"""
    bp = Blueprint("pages", __name__)

    # 应用 csrf_exempt 装饰器（如果提供）
    login_view = pages_controller.login
    csrf_token_view = pages_controller.get_csrf_token
    api_logout_view = pages_controller.api_logout
    if csrf_exempt is not None:
        login_view = csrf_exempt(login_view)
        csrf_token_view = csrf_exempt(csrf_token_view)
        api_logout_view = csrf_exempt(api_logout_view)

    bp.add_url_rule("/login", view_func=login_view, methods=["GET", "POST"])
    bp.add_url_rule("/logout", view_func=pages_controller.logout, methods=["GET", "POST"])
    bp.add_url_rule("/", view_func=pages_controller.index, methods=["GET"])
    bp.add_url_rule("/mailbox", view_func=pages_controller.removed_mailbox_compat, methods=["GET"])
    bp.add_url_rule("/ops", view_func=pages_controller.admin_index, methods=["GET"])
    bp.add_url_rule("/img/<path:filename>", view_func=pages_controller.image_asset, methods=["GET"])
    bp.add_url_rule("/favicon.ico", view_func=pages_controller.favicon, methods=["GET"])
    bp.add_url_rule("/api/csrf-token", view_func=csrf_token_view, methods=["GET"])
    # SPA / Ant Design Pro 鉴权适配
    bp.add_url_rule("/api/auth/current-user", view_func=pages_controller.api_current_user, methods=["GET"])
    bp.add_url_rule("/api/auth/logout", view_func=api_logout_view, methods=["POST", "GET"])
    return bp
