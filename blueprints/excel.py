"""Routy pro upload, download a prohlížení Excel souborů."""

from flask import Blueprint, flash, g, redirect, render_template, request, send_file, send_from_directory, url_for

from performance_optimizations import invalidate_excel_status_cache
from services.excel_view_service import get_excel_editor_context, get_excel_viewer_context
from services.upload_service import (
    confirm_overwrite,
    normalize_upload_filename,
    read_and_validate_excel,
    save_uploaded_file,
    store_temp_upload,
)
from utils.logger import setup_logger

logger = setup_logger("excel_routes")

excel_bp = Blueprint("excel", __name__)


@excel_bp.route("/download")
def download_file():
    """Stáhne aktuálně aktivní Excel soubor."""
    try:
        return send_file(g.excel_manager.get_active_file_path(), as_attachment=True)
    except FileNotFoundError as exc:
        logger.error("Chyba při stahování souboru: %s", exc, exc_info=True)
        flash("Chyba při stahování souboru.", "error")
        return redirect(url_for("main.index"))


@excel_bp.route("/download/<path:filename>")
def download_specific_file(filename):
    """Stáhne specifický soubor z upload složky."""
    try:
        return send_from_directory(g.excel_manager.base_path, filename, as_attachment=True)
    except FileNotFoundError:
        logger.error("Pokus o stažení neexistujícího souboru: %s", filename, exc_info=True)
        flash("Soubor nenalezen.", "error")
        return redirect(url_for("excel.excel_viewer"))


@excel_bp.route("/upload", methods=["GET", "POST"])
def upload_file():
    """Nahrání Excel souborů s kontrolou přepsání existujících souborů."""
    if request.method != "POST":
        return redirect(url_for("main.index"))

    if "overwrite" in request.form and "filename" in request.form:
        flash("Pro dokončení přepsání prosím znovu vyberte soubor.", "info")
        return redirect(url_for("main.index"))

    if "file" not in request.files:
        flash("Nebyl vybrán žádný soubor.", "error")
        return redirect(url_for("main.index"))

    file = request.files["file"]
    if file.filename == "":
        flash("Nebyl vybrán žádný soubor.", "error")
        return redirect(url_for("main.index"))

    try:
        filename = normalize_upload_filename(file.filename)
        category = request.form.get("category", "Ostatní")
        file_path = g.excel_manager.base_path / filename
        force_overwrite = request.form.get("force_overwrite") == "true"

        if file_path.exists() and not force_overwrite:
            temp_filename = store_temp_upload(filename, read_and_validate_excel(file))
            return render_template("upload_confirm.html", filename=filename, temp_filename=temp_filename)

        file_bytes = read_and_validate_excel(file)
        save_uploaded_file(filename, file_bytes)
        g.excel_manager.set_category(filename, category)
        invalidate_excel_status_cache()
        flash(f"Soubor '{filename}' byl úspěšně nahrán.", "success")
        logger.info("Nahrán soubor: %s", filename)
    except ValueError as exc:
        flash(str(exc), "error")
    except Exception as exc:
        logger.error("Chyba při nahrávání souboru: %s", exc, exc_info=True)
        flash("Chyba při nahrávání souboru.", "error")

    return redirect(url_for("main.index"))


@excel_bp.route("/upload/confirm", methods=["POST"])
def upload_confirm():
    """Potvrzení přepsání existujícího souboru."""
    temp_filename = request.form.get("temp_filename")
    filename = request.form.get("filename")

    try:
        confirm_overwrite(temp_filename, filename)
        invalidate_excel_status_cache()
        flash(f"Soubor '{filename}' byl úspěšně přepsán.", "success")
        logger.info("Přepsán soubor: %s", filename)
    except (ValueError, FileNotFoundError) as exc:
        flash(str(exc), "error")
    except Exception as exc:
        logger.error("Chyba při přepisování souboru: %s", exc, exc_info=True)
        flash("Chyba při přepisování souboru.", "error")
        try:
            if temp_filename:
                temp_path = g.excel_manager.base_path / temp_filename
                if temp_path.exists():
                    temp_path.unlink()
        except Exception:
            pass

    return redirect(url_for("main.index"))


@excel_bp.route("/excel_viewer", methods=["GET"])
def excel_viewer():
    """Read-only náhled omezeného počtu řádků aktivního Excelu."""
    try:
        context = get_excel_viewer_context(
            g.excel_manager,
            requested_file=request.args.get("file"),
            active_sheet_name=request.args.get("sheet"),
            selected_category=request.args.get("category"),
        )
    except FileNotFoundError as exc:
        flash(str(exc), "error")
        return redirect(url_for("main.index"))

    if not context["excel_files"]:
        flash("Nenalezeny žádné Excel soubory pro danou kategorii.", "warning")
    elif context["selected_file"] and not context["sheet_names"]:
        flash("Vybraný soubor nemá žádné listy.", "warning")

    return render_template("excel_viewer.html", **context)


@excel_bp.route("/excel_editor", methods=["GET", "POST"])
def excel_editor():
    """Editable náhled Excel souborů s možností úprav přímo v prohlížeči."""
    if request.method == "POST":
        try:
            file_name = request.form.get("file")
            sheet_name = request.form.get("sheet")
            row_str = request.form.get("row")
            col_str = request.form.get("col")
            value = request.form.get("value")

            if not file_name or not sheet_name or not row_str or not col_str:
                flash("Chybí název souboru, listu nebo pozice buňky.", "error")
                return redirect(url_for("excel.excel_editor"))

            row = int(row_str)
            col = int(col_str)

            if g.excel_manager.update_cell(file_name, sheet_name, row, col, value):
                flash("Buňka byla úspěšně aktualizována.", "success")
            else:
                flash("Nepodařilo se aktualizovat buňku.", "error")

            return redirect(url_for("excel.excel_editor", file=file_name, sheet=sheet_name))
        except (TypeError, ValueError):
            flash("Neplatná pozice buňky.", "error")
            return redirect(url_for("excel.excel_editor"))
        except Exception as exc:
            logger.error("Chyba v excel_editor POST: %s", exc, exc_info=True)
            flash(f"Chyba při ukládání: {exc}", "error")
            return redirect(url_for("excel.excel_editor"))

    try:
        context = get_excel_editor_context(
            g.excel_manager.active_filename,
            requested_file=request.args.get("file"),
            active_sheet_name=request.args.get("sheet"),
        )
    except FileNotFoundError as exc:
        flash(str(exc), "error")
        return redirect(url_for("main.index"))

    if not context["excel_files"]:
        flash("Nenalezeny žádné Excel soubory.", "warning")
    elif context["selected_file"] and not context["sheet_names"]:
        flash("Vybraný soubor nemá žádné listy.", "warning")

    return render_template("excel_editor.html", **context)
