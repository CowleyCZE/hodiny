"""Routy pro běžné aplikační nastavení."""

from flask import Blueprint, flash, render_template, request, session

from performance_optimizations import invalidate_user_settings_cache
from services.settings_service import save_app_settings

settings_bp = Blueprint("settings", __name__)


@settings_bp.route("/settings", methods=["GET", "POST"])
def settings_page():
    """Nastavení výchozí pracovní doby a základních údajů projektu."""
    if request.method == "POST":
        try:
            settings_to_save = session["settings"].copy()
            settings_to_save.update(
                {
                    "start_time": request.form["start_time"],
                    "end_time": request.form["end_time"],
                    "lunch_duration": float(request.form["lunch_duration"].replace(",", ".")),
                }
            )
            settings_to_save["project_info"] = {
                "name": request.form.get("project_name", "").strip(),
                "start_date": request.form.get("start_date", "").strip(),
                "end_date": request.form.get("end_date", "").strip(),
            }
            if not save_app_settings(settings_to_save):
                raise IOError("Nepodařilo se uložit nastavení.")
            session["settings"] = settings_to_save
            invalidate_user_settings_cache()
            flash("Nastavení bylo úspěšně uloženo.", "success")
        except (ValueError, IOError) as exc:
            flash(str(exc), "error")

    return render_template("settings_page.html", settings=session.get("settings", {}))
