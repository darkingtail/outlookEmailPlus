from __future__ import annotations

from flask import Blueprint, jsonify

from outlook_web.controllers import customer_mail


def create_blueprint(csrf_exempt):
    bp = Blueprint("customer_mail", __name__)

    @bp.after_request
    def private_response(response):
        response.headers["Cache-Control"] = "no-store"
        response.headers["Referrer-Policy"] = "no-referrer"
        return response

    bp.add_url_rule(
        "/api/customer/messages",
        view_func=csrf_exempt(customer_mail.messages),
        methods=["POST"],
    )
    bp.add_url_rule(
        "/api/customer/messages/detail",
        view_func=csrf_exempt(customer_mail.message_detail),
        methods=["POST"],
    )

    @bp.route("/api/customer/messages/detail", methods=["GET"])
    def detail_post_only():
        # Prevent the SPA catch-all from turning an incorrect API GET into HTML/200.
        response = jsonify({"success": False, "message": "请使用 POST 请求"})
        response.status_code = 405
        response.headers["Allow"] = "POST, OPTIONS"
        return response

    return bp
