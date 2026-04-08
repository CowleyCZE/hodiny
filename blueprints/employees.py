"""Routy pro správu zaměstnanců."""

from flask import Blueprint, flash, g, render_template, request

from performance_optimizations import invalidate_employee_stats_cache

employees_bp = Blueprint("employees", __name__)


@employees_bp.route("/zamestnanci", methods=["GET", "POST"])
def manage_employees():
    """Správa zaměstnanců (přidání, výběr pro zapisování, editace, mazání)."""
    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "add":
                g.employee_manager.pridat_zamestnance(request.form.get("name", "").strip())
            elif action == "select":
                name = request.form.get("employee_name", "")
                if name in g.employee_manager.vybrani_zamestnanci:
                    g.employee_manager.odebrat_vybraneho_zamestnance(name)
                else:
                    g.employee_manager.pridat_vybraneho_zamestnance(name)
            elif action == "edit":
                old_name = request.form.get("old_name", "").strip()
                new_name = request.form.get("new_name", "").strip()
                g.employee_manager.upravit_zamestnance_podle_jmena(old_name, new_name)
            elif action == "delete":
                g.employee_manager.smazat_zamestnance_podle_jmena(request.form.get("employee_name", ""))
            invalidate_employee_stats_cache()
        except ValueError as exc:
            flash(str(exc), "error")
    return render_template("employees.html", employees=g.employee_manager.get_all_employees())
