"""Bootstrap Flask aplikace a request lifecycle pro projekt Hodiny."""

import datetime as dt

from flask import Flask, g, redirect, request, session, url_for

from api_endpoints import api_bp
from blueprints.auth import auth_bp
from blueprints.configuration import configuration_bp
from blueprints.employees import employees_bp
from blueprints.excel import excel_bp
from blueprints.main import main_bp
from blueprints.reports import reports_bp
from blueprints.settings import settings_bp
from config import Config
from employee_management import EmployeeManager
from excel_manager import ExcelManager
from hodiny2025_manager import Hodiny2025Manager
from performance_optimizations import cleanup_old_data, initialize_performance_optimizations
from security import csrf_token, validate_csrf
from services.database import Database
from services.settings_service import load_app_settings, save_app_settings
from utils.logger import setup_logger
from zalohy_manager import ZalohyManager

logger = setup_logger("app")

app = Flask(__name__)
app.secret_key = Config.SECRET_KEY
Config.init_app(app)
app.jinja_env.globals["csrf_token"] = csrf_token


@app.before_request
def security_before_request():
    validate_csrf()
    if Config.ADMIN_USERNAME and Config.ADMIN_PASSWORD and request.endpoint != "health_check":
        if not session.get("authenticated"):
            auth = request.authorization
            if not auth or not (auth.username == Config.ADMIN_USERNAME and auth.password == Config.ADMIN_PASSWORD):
                return ("Přístup odepřen", 401, {"WWW-Authenticate": 'Basic realm="hodiny"'})


app.register_blueprint(api_bp)
app.register_blueprint(auth_bp)
app.register_blueprint(configuration_bp)
app.register_blueprint(employees_bp)
app.register_blueprint(excel_bp)
app.register_blueprint(main_bp)
app.register_blueprint(reports_bp)
app.register_blueprint(settings_bp)

initialize_performance_optimizations()
database = Database(Config.DATABASE_PATH)


@app.after_request
def security_headers(response):
    """Základní hlavičky proti běžným webovým útokům."""
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    response.headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
    return response


def initialize_archived_state():
    """Run one archive check at process startup, not once per HTTP request."""
    settings = load_app_settings()
    manager = ExcelManager(Config.EXCEL_BASE_PATH, hodiny2025_manager=Hodiny2025Manager(Config.EXCEL_BASE_PATH))
    current_week = dt.datetime.now().isocalendar().week
    if manager.archive_if_needed(current_week, settings):
        save_app_settings(settings)
    manager.close_cached_workbooks()


initialize_archived_state()


@app.before_request
def before_request():
    """Před každým requestem připraví managery a synchronizuje runtime nastavení."""
    g.database = database
    g.current_user = database.get_user(session["user_id"]) if session.get("user_id") else None
    if Config.AUTH_REQUIRED and request.endpoint != "static" and not request.path.startswith("/auth/"):
        if request.path == "/api/v1/health":
            return None
        if database.user_count() == 0:
            return redirect(url_for("auth.setup"))
        if g.current_user is None:
            return (
                redirect(url_for("auth.login", next=request.path))
                if not request.path.startswith("/api/")
                else ({"success": False, "error": "Přihlášení je vyžadováno."}, 401)
            )
    session["settings"] = load_app_settings()
    g.employee_manager = EmployeeManager(
        Config.DATA_PATH,
        preferred_employee_name=session["settings"].get("preferred_employee_name", ""),
    )
    g.hodiny2025_manager = Hodiny2025Manager(Config.EXCEL_BASE_PATH)
    g.excel_manager = ExcelManager(Config.EXCEL_BASE_PATH, hodiny2025_manager=g.hodiny2025_manager)
    g.zalohy_manager = ZalohyManager(Config.EXCEL_BASE_PATH)
    g.excel_manager.update_project_info(
        session["settings"].get("project_info", {}).get("name", ""),
        session["settings"].get("project_info", {}).get("start_date", ""),
        session["settings"].get("project_info", {}).get("end_date", ""),
    )

    cleanup_old_data()


@app.teardown_request
def teardown_request(_exception=None):
    """Uzavře případné otevřené workbooky po dokončení requestu."""
    if hasattr(g, "excel_manager") and g.excel_manager:
        g.excel_manager.close_cached_workbooks()

@app.route("/webhook/github", methods=["POST"])
def github_webhook():
    """GitHub webhook pro automatický deployment"""
    import subprocess
    import os
    from flask import request
    
    # Základní bezpečnost - ověř GitHub secret (pokud chceš)
    # signature = request.headers.get('X-Hub-Signature-256', '')
    # Zatím bez ověření pro jednoduchost
    
    event = request.headers.get('X-GitHub-Event')
    
    if event == 'push':
        try:
            # Spusť git pull v adresáři aplikace
            repo_path = "/home/Cowley/hodiny"
            
            result = subprocess.run(
                ["git", "-C", repo_path, "fetch", "origin"],
                capture_output=True,
                text=True,
                timeout=30
            )
            
            result = subprocess.run(
                ["git", "-C", repo_path, "reset", "--hard", "origin/main"],
                capture_output=True,
                text=True,
                timeout=30
            )
            
            return {"status": "success", "message": "Deployment completed"}, 200
        
        except Exception as e:
            print(f"Webhook error: {e}")
            return {"status": "error", "message": str(e)}, 500
    
    return {"status": "ignored"}, 200

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
