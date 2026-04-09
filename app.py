"""Bootstrap Flask aplikace a request lifecycle pro projekt Hodiny."""

import datetime as dt
import random

from flask import Flask, flash, g, session

from api_endpoints import api_bp
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
from services.settings_service import load_app_settings, save_app_settings
from utils.logger import setup_logger
from zalohy_manager import ZalohyManager

logger = setup_logger("app")

app = Flask(__name__)
app.secret_key = Config.SECRET_KEY
Config.init_app(app)

app.register_blueprint(api_bp)
app.register_blueprint(configuration_bp)
app.register_blueprint(employees_bp)
app.register_blueprint(excel_bp)
app.register_blueprint(main_bp)
app.register_blueprint(reports_bp)
app.register_blueprint(settings_bp)

initialize_performance_optimizations()


@app.before_request
def before_request():
    """Před každým requestem připraví managery a synchronizuje runtime nastavení."""
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

    if random.randint(1, 100) == 1:
        cleanup_old_data()

    current_week = dt.datetime.now().isocalendar().week
    if g.excel_manager.archive_if_needed(current_week, session["settings"]):
        save_app_settings(session["settings"])
        flash(f"Týden {session['settings']['last_archived_week'] - 1} byl archivován.", "info")


@app.teardown_request
def teardown_request(_exception=None):
    """Uzavře případné otevřené workbooky po dokončení requestu."""
    if hasattr(g, "excel_manager") and g.excel_manager:
        g.excel_manager.close_cached_workbooks()


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
