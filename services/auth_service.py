"""Autentizace, role a řízení přístupu."""

from functools import wraps

from flask import g, jsonify, redirect, request, session, url_for


def current_user():
    return getattr(g, "current_user", None)


def login_user(user):
    session.clear()
    session["user_id"] = user["id"]
    session["username"] = user["username"]
    session["role"] = user["role"]
    session.permanent = True


def logout_user():
    session.clear()


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if current_user() is None:
            if request.path.startswith("/api/"):
                return jsonify({"success": False, "error": "Přihlášení je vyžadováno."}), 401
            return redirect(url_for("auth.login", next=request.path))
        return view(*args, **kwargs)

    return wrapped


def role_required(*roles):
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            user = current_user()
            if user is None:
                return login_required(view)(*args, **kwargs)
            if user["role"] not in roles:
                if request.path.startswith("/api/"):
                    return jsonify({"success": False, "error": "Nedostatečné oprávnění."}), 403
                return "Nedostatečné oprávnění.", 403
            return view(*args, **kwargs)

        return wrapped

    return decorator
