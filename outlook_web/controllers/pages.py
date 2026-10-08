from __future__ import annotations

from pathlib import Path
from typing import Any

from flask import abort, g, jsonify, redirect, render_template, request, send_from_directory, session, url_for

from outlook_web import spa as spa_support
from outlook_web.errors import build_error_payload
from outlook_web.repositories import settings as settings_repo
from outlook_web.security.auth import (
    check_rate_limit,
    get_client_ip,
    login_required,
    record_login_failure,
    reset_login_attempts,
)
from outlook_web.security.crypto import verify_password

# ==================== 页面路由 ====================


def login() -> Any:
    """登录页面"""
    if request.method == "POST":
        try:
            # 获取客户端 IP
            client_ip = get_client_ip()

            # 检查速率限制
            allowed, remaining_time = check_rate_limit(client_ip)
            if not allowed:
                trace_id_value = None
                try:
                    trace_id_value = getattr(g, "trace_id", None)
                except Exception:
                    trace_id_value = None
                error_payload = build_error_payload(
                    code="LOGIN_RATE_LIMITED",
                    message=f"登录失败次数过多，请在 {remaining_time} 秒后重试",
                    err_type="RateLimitError",
                    status=429,
                    details=f"ip={client_ip}",
                    trace_id=trace_id_value,
                )
                return jsonify({"success": False, "error": error_payload}), 429

            data = request.json if request.is_json else request.form
            password = data.get("password", "")

            # 从数据库获取密码哈希
            stored_password = settings_repo.get_login_password()

            # 验证密码
            if verify_password(password, stored_password):
                # 登录成功，重置失败记录
                reset_login_attempts(client_ip)
                session["logged_in"] = True
                session.permanent = True
                session.modified = True  # 确保 Flask-Session 保存 session
                return jsonify({"success": True, "message": "登录成功"})
            else:
                # 登录失败，记录失败次数
                record_login_failure(client_ip)
                trace_id_value = None
                try:
                    trace_id_value = getattr(g, "trace_id", None)
                except Exception:
                    trace_id_value = None
                error_payload = build_error_payload(
                    code="LOGIN_INVALID_PASSWORD",
                    message="密码错误",
                    message_en="Invalid password",
                    err_type="AuthError",
                    status=401,
                    details=f"ip={client_ip}",
                    trace_id=trace_id_value,
                )
                return jsonify({"success": False, "error": error_payload}), 401
        except Exception as e:
            trace_id_value = None
            try:
                trace_id_value = getattr(g, "trace_id", None)
            except Exception:
                trace_id_value = None
            from flask import current_app

            try:
                current_app.logger.exception("Login error trace_id=%s", trace_id_value or "unknown")
            except Exception:
                pass
            error_payload = build_error_payload(
                code="LOGIN_FAILED",
                message="登录处理失败",
                err_type="AuthError",
                status=500,
                details=str(e),
                trace_id=trace_id_value,
            )
            return jsonify({"success": False, "error": error_payload}), 500

    # GET 请求返回登录页面（SPA 模式下由前端路由处理）
    if spa_support.spa_enabled():
        return spa_support.send_spa_index()
    return render_template("login.html")


def logout() -> Any:
    """退出登录。

    兼容旧版页面跳转与新前端 SPA：
    - 默认 GET /logout → 重定向到登录页
    - Accept: application/json 或 /api/* 路径 → JSON
    """
    session.pop("logged_in", None)
    wants_json = (
        request.is_json
        or request.path.startswith("/api/")
        or "application/json" in (request.headers.get("Accept") or "")
        or request.headers.get("X-Requested-With") == "XMLHttpRequest"
    )
    if wants_json:
        return jsonify({"success": True, "message": "已退出登录"})
    return redirect(url_for("pages.login"))


def api_current_user() -> Any:
    """当前登录用户（供 Ant Design Pro / SPA 使用）。

    未登录返回 401 + need_login，已登录返回统一 data 结构。
    本系统是单密码管理端，无多用户模型，固定返回管理员身份。
    """
    is_logged_in = bool(session.get("logged_in") or session.get("user_id"))
    if not is_logged_in:
        trace_id_value = None
        try:
            trace_id_value = getattr(g, "trace_id", None)
        except Exception:
            trace_id_value = None
        error_payload = build_error_payload(
            code="AUTH_REQUIRED",
            message="请先登录",
            err_type="AuthError",
            status=401,
            details="need_login",
            trace_id=trace_id_value,
        )
        return jsonify({"success": False, "error": error_payload, "need_login": True}), 401

    return jsonify(
        {
            "success": True,
            "data": {
                "name": "管理员",
                "userid": "admin",
                "access": "admin",
                "avatar": "/img/ico.png",
                "title": "Outlook 邮件管理",
            },
        }
    )


def api_logout() -> Any:
    """SPA 退出登录（JSON）。"""
    session.pop("logged_in", None)
    return jsonify({"success": True, "message": "已退出登录"})


def image_asset(filename: str) -> Any:
    """提供仓库 img/ 目录中的静态图片资源。

    Docker 部署镜像通过 .dockerignore 排除了仓库 img/ 目录，
    仓库内文件缺失时回退到 SPA 构建产物（ant-design-pro/public/img/），
    保证新前端的 logo 等资源始终可加载。
    """
    img_dir = Path(__file__).resolve().parents[2] / "img"
    try:
        resolved = (img_dir / filename).resolve()
        resolved.relative_to(img_dir.resolve())
    except (OSError, ValueError):
        resolved = None
    if resolved is not None and resolved.is_file():
        return send_from_directory(str(img_dir), filename)
    spa_response = spa_support.send_spa_file(f"img/{filename}")
    if spa_response is not None:
        return spa_response
    abort(404)


def favicon() -> Any:
    """站点 favicon，复用 img/ico.png；仓库 img/ 缺失时回退 SPA 构建产物。"""
    return image_asset("ico.png")


def index() -> Any:
    """客户邮箱查询页。"""
    return render_template("customer_mail.html")


def removed_mailbox_compat() -> Any:
    """顶层旧客户入口已移除，避免被 SPA 回退误当成有效页面。"""
    abort(404)


def admin_index() -> Any:
    """管理员 SPA 入口。

    页面外壳允许匿名加载，以便前端显示登录页；管理员数据接口仍由
    后端登录鉴权保护。
    """
    if spa_support.spa_enabled():
        return spa_support.send_spa_index()
    return render_template("index.html")


def get_csrf_token() -> Any:
    """获取CSRF Token"""
    from outlook_web.security.csrf import CSRF_AVAILABLE, generate_csrf

    if CSRF_AVAILABLE:
        token = generate_csrf()
        return jsonify({"csrf_token": token})
    else:
        return jsonify({"csrf_token": None, "csrf_disabled": True})
