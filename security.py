"""Shared security helpers for filesystem and request protection."""
from pathlib import Path
import secrets

from flask import abort, current_app, request, session


def safe_excel_path(base_path, filename):
    """Return a path below *base_path*; reject traversal and non-XLSX files."""
    if not filename or Path(filename).name != filename or not filename.lower().endswith(".xlsx"):
        raise ValueError("Neplatný název Excel souboru.")
    base = Path(base_path).resolve()
    target = (base / filename).resolve()
    if target.parent != base:
        raise ValueError("Neplatná cesta k souboru.")
    return target


def csrf_token():
    token = session.get("csrf_token")
    if not token:
        token = secrets.token_urlsafe(32)
        session["csrf_token"] = token
    return token


def validate_csrf():
    """Validate CSRF for browser form requests; JSON APIs use their API boundary."""
    if request.method in {"GET", "HEAD", "OPTIONS"} or request.is_json or current_app.testing:
        return
    expected = session.get("csrf_token")
    supplied = request.form.get("csrf_token") or request.headers.get("X-CSRFToken")
    if not expected or not supplied or not secrets.compare_digest(expected, supplied):
        abort(400, description="Neplatný CSRF token.")
