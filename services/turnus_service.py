"""Služby pro správu životního cyklu a validace turnusů (výjezdů / montáží)."""

import datetime as dt
import json
import re
import shutil
from pathlib import Path

from config import Config
from services.phase2_services import write_audit
from services.settings_service import load_app_settings, save_app_settings
from utils.logger import setup_logger

logger = setup_logger("turnus_service")


def _get_turnus_file_path(filepath=None):
    if filepath is not None:
        return Path(filepath)
    return Config.DATA_PATH / "turnusy.json"


def _format_cz_date(date_str):
    """Převede YYYY-MM-DD na DD.MM.YYYY pro české chybové hlášky."""
    try:
        parts = date_str.split("-")
        if len(parts) == 3:
            return f"{int(parts[2])}.{int(parts[1])}.{parts[0]}"
    except Exception:
        pass
    return date_str


def _create_default_turnus_from_settings(settings):
    project_info = settings.get("project_info", {})
    name = project_info.get("name", "").strip() or "Blankendaal"
    start_date = project_info.get("start_date", "").strip() or "2020-01-01"
    end_date = project_info.get("end_date", "").strip() or None
    firma = project_info.get("firma", "").strip()
    mesto = project_info.get("mesto", "").strip()

    turnus_id = f"turnus_{re.sub(r'[^a-zA-Z0-9_]', '_', name.lower())}_{start_date.replace('-', '')}"
    return {
        "active_turnus_id": turnus_id,
        "turnusy": [
            {
                "id": turnus_id,
                "name": name,
                "firma": firma,
                "mesto": mesto,
                "start_date": start_date,
                "end_date": end_date,
                "status": "active",
                "created_at": dt.datetime.now().isoformat(),
                "closed_at": None,
                "archive_file": None,
            }
        ],
    }


def load_turnus_data(filepath=None):
    """Načte data turnusů z JSON souboru nebo je inicializuje ze settings.json."""
    path = _get_turnus_file_path(filepath)
    if not path.exists():
        settings = load_app_settings()
        default_data = _create_default_turnus_from_settings(settings)
        save_turnus_data(default_data, filepath=path)
        return default_data

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            if not isinstance(data, dict):
                raise ValueError("Neplatný formát dat turnusů.")
            if "turnusy" not in data:
                data["turnusy"] = []
            if "active_turnus_id" not in data:
                data["active_turnus_id"] = None
            return data
    except (json.JSONDecodeError, IOError, ValueError) as exc:
        logger.error("Chyba při čtení %s: %s", path, exc, exc_info=True)
        return {"active_turnus_id": None, "turnusy": []}


def save_turnus_data(data, filepath=None):
    """Uloží normalizovaná data turnusů do JSON souboru."""
    path = _get_turnus_file_path(filepath)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        return True
    except (IOError, TypeError) as exc:
        logger.error("Chyba při ukládání %s: %s", path, exc, exc_info=True)
        return False


def get_all_turnusy(filepath=None):
    """Vrátí seznam všech turnusů (seřazený od nejnovějšího)."""
    data = load_turnus_data(filepath)
    turnusy = list(data.get("turnusy", []))
    # Seřadit podle start_date sestupně
    turnusy.sort(key=lambda t: t.get("start_date", ""), reverse=True)
    return turnusy


def get_active_turnus(filepath=None):
    """Vrátí aktuálně běžící aktivní turnus nebo None."""
    data = load_turnus_data(filepath)
    active_id = data.get("active_turnus_id")
    if not active_id:
        # Fallback: zkontroluj, zda v turnusy není položka se status == 'active'
        for t in data.get("turnusy", []):
            if t.get("status") == "active":
                return t
        return None

    for t in data.get("turnusy", []):
        if t.get("id") == active_id and t.get("status") == "active":
            return t
    return None


def get_turnus_by_id(turnus_id, filepath=None):
    """Vrátí turnus dle ID."""
    data = load_turnus_data(filepath)
    for t in data.get("turnusy", []):
        if t.get("id") == turnus_id:
            return t
    return None


def create_turnus(name, start_date, end_date=None, firma="", mesto="", filepath=None):
    """Založí nový aktivní turnus a aktualizuje project_info v settings."""
    name = (name or "").strip()
    if not name:
        raise ValueError("Název turnusu je povinný.")

    start_date = (start_date or "").strip()
    if not start_date:
        raise ValueError("Datum začátku turnusu je povinné.")

    try:
        start_dt = dt.datetime.strptime(start_date, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError("Datum začátku musí mít formát RRRR-MM-DD.") from exc

    clean_end_date = None
    if end_date:
        end_date = end_date.strip()
        if end_date:
            try:
                end_dt = dt.datetime.strptime(end_date, "%Y-%m-%d").date()
                if end_dt < start_dt:
                    raise ValueError("Datum ukončení nesmí být před datem začátku.")
                clean_end_date = end_date
            except ValueError as exc:
                if "Datum ukončení nesmí" in str(exc):
                    raise
                raise ValueError("Datum ukončení musí mít formát RRRR-MM-DD.") from exc

    data = load_turnus_data(filepath)
    active = get_active_turnus(filepath)
    if active:
        raise ValueError(
            f"Nelze vytvořit nový turnus: Stále běží aktivní turnus '{active['name']}'. "
            "Před založením nového jej nejprve ukončete."
        )

    slug = re.sub(r"[^a-zA-Z0-9_]", "_", name.lower())
    turnus_id = f"turnus_{slug}_{start_date.replace('-', '')}_{dt.datetime.now().strftime('%H%M%S')}"

    new_turnus = {
        "id": turnus_id,
        "name": name,
        "firma": (firma or "").strip(),
        "mesto": (mesto or "").strip(),
        "start_date": start_date,
        "end_date": clean_end_date,
        "status": "active",
        "created_at": dt.datetime.now().isoformat(),
        "closed_at": None,
        "archive_file": None,
    }

    data.setdefault("turnusy", []).append(new_turnus)
    data["active_turnus_id"] = turnus_id
    save_turnus_data(data, filepath)

    # Zrcadlit do settings.json pro Excel list a další návaznosti
    if filepath is None:
        try:
            settings = load_app_settings()
            settings["project_info"] = {
                "name": name,
                "start_date": start_date,
                "end_date": clean_end_date or "",
                "firma": (firma or "").strip(),
                "mesto": (mesto or "").strip(),
            }
            save_app_settings(settings)
        except Exception as exc:
            logger.warning("Synchronizace turnusu do settings.json selhala: %s", exc)

    write_audit(Config.AUDIT_LOG_PATH, "TURNUS_CREATED", {"id": turnus_id, "name": name, "start": start_date})
    return new_turnus


def close_turnus(turnus_id=None, end_date=None, filepath=None, excel_path=None, backup_dir=None):
    """Ukončí turnus k zadanému datu, nastaví status=closed a vytvoří archivní kopii."""
    data = load_turnus_data(filepath)
    target = None

    if turnus_id:
        for t in data.get("turnusy", []):
            if t.get("id") == turnus_id:
                target = t
                break
    else:
        active_id = data.get("active_turnus_id")
        for t in data.get("turnusy", []):
            if t.get("id") == active_id:
                target = t
                break

    if not target:
        raise ValueError("Nebyl nalezen žádný aktivní turnus k ukončení.")

    if target.get("status") == "closed":
        raise ValueError(f"Turnus '{target['name']}' je již ukončen.")

    if not end_date:
        end_date = dt.date.today().isoformat()
    else:
        end_date = end_date.strip()

    try:
        end_dt = dt.datetime.strptime(end_date, "%Y-%m-%d").date()
        start_dt = dt.datetime.strptime(target["start_date"], "%Y-%m-%d").date()
        if end_dt < start_dt:
            raise ValueError(
                f"Datum ukončení ({_format_cz_date(end_date)}) nesmí předcházet začátku turnusu "
                f"({_format_cz_date(target['start_date'])})."
            )
    except ValueError as exc:
        if "Datum ukončení" in str(exc):
            raise
        raise ValueError("Datum ukončení musí mít formát RRRR-MM-DD.") from exc

    # Vytvoření archivní kopie (Varianta A)
    archive_rel_path = None
    source_excel = Path(excel_path or (Config.EXCEL_BASE_PATH / Config.EXCEL_TEMPLATE_NAME))
    target_backup_dir = Path(backup_dir or Config.BACKUP_PATH)

    if source_excel.exists():
        try:
            target_backup_dir.mkdir(parents=True, exist_ok=True)
            safe_name = re.sub(r"[^\w\-_]", "_", target["name"])
            archive_filename = f"Turnus_{safe_name}_{target['start_date']}_{end_date}.xlsx"
            archive_full_path = target_backup_dir / archive_filename
            shutil.copy2(source_excel, archive_full_path)
            archive_rel_path = str(archive_full_path.relative_to(Config.BASE_DIR)) if archive_full_path.is_relative_to(Config.BASE_DIR) else str(archive_full_path)
            logger.info("Vytvořena archivní kopie turnusu: %s", archive_full_path)
        except Exception as exc:
            logger.warning("Nepodařilo se vytvořit archivní kopii Excelu: %s", exc)

    target["status"] = "closed"
    target["end_date"] = end_date
    target["closed_at"] = dt.datetime.now().isoformat()
    target["archive_file"] = archive_rel_path

    if data.get("active_turnus_id") == target["id"]:
        data["active_turnus_id"] = None

    save_turnus_data(data, filepath)

    # Aktualizace settings
    if filepath is None:
        try:
            settings = load_app_settings()
            if settings.get("project_info", {}).get("name") == target["name"]:
                settings["project_info"]["end_date"] = end_date
                save_app_settings(settings)
        except Exception as exc:
            logger.warning("Aktualizace settings.json při ukončení turnusu selhala: %s", exc)

    write_audit(
        Config.AUDIT_LOG_PATH,
        "TURNUS_CLOSED",
        {"id": target["id"], "name": target["name"], "end_date": end_date, "archive": archive_rel_path},
    )
    return target


def validate_date_for_entry(date_str, filepath=None):
    """Zkontroluje, zda datum zápisu spadá do platného období turnusu.

    Vrací tuple (is_valid: bool, error_message: str, turnus: dict | None).
    Pravidla:
    1. Datum spadá do aktivního turnusu (start_date <= date, a pokud má end_date, pak date <= end_date) -> POVOLENO.
    2. Datum spadá do uzavřeného turnusu (start_date <= date <= end_date) -> POVOLENO (úprava/doplnění historie).
    3. Pokud pro dané datum není nalezen žádný turnus -> ZAMÍTNUTO s vysvětlující hláškou.
    """
    if not date_str:
        return False, "Chybí datum záznamu.", None

    try:
        entry_date = dt.datetime.strptime(date_str, "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return False, f"Neplatný formát data '{date_str}'. Očekává se RRRR-MM-DD.", None

    formatted_entry = _format_cz_date(date_str)
    turnusy = get_all_turnusy(filepath)

    if not turnusy:
        return (
            False,
            f"Záznam pro {formatted_entry} nelze uložit: V aplikaci není vytvořen žádný turnus. "
            "Nejprve prosím založte turnus v záložce Turnus.",
            None,
        )

    active_turnus = get_active_turnus(filepath)

    # 1. Kontrola aktivního turnusu
    if active_turnus:
        try:
            act_start = dt.datetime.strptime(active_turnus["start_date"], "%Y-%m-%d").date()
        except ValueError:
            act_start = None

        act_end = None
        if active_turnus.get("end_date"):
            try:
                act_end = dt.datetime.strptime(active_turnus["end_date"], "%Y-%m-%d").date()
            except ValueError:
                act_end = None

        if act_start and entry_date >= act_start:
            if act_end is None or entry_date <= act_end:
                return True, "", active_turnus

    # 2. Kontrola všech turnusů (včetně uzavřených – doplnění/úprava zapomenutých dnů)
    for t in turnusy:
        try:
            t_start = dt.datetime.strptime(t["start_date"], "%Y-%m-%d").date()
        except ValueError:
            continue

        t_end = None
        if t.get("end_date"):
            try:
                t_end = dt.datetime.strptime(t["end_date"], "%Y-%m-%d").date()
            except ValueError:
                t_end = None

        if t_end is not None:
            if t_start <= entry_date <= t_end:
                return True, "", t
        elif t.get("status") == "active" and t_start <= entry_date:
            return True, "", t

    # 3. Zápis byl zamítnut – zformulujeme přesný důvod
    # A) Zkouší zapsat do budoucnosti po ukončení aktivního turnusu s nastaveným end_date
    if active_turnus and active_turnus.get("end_date"):
        act_end_dt = dt.datetime.strptime(active_turnus["end_date"], "%Y-%m-%d").date()
        if entry_date > act_end_dt:
            return (
                False,
                f"Záznam pro {formatted_entry} nelze uložit: Datum je po plánovaném ukončení turnusu "
                f"'{active_turnus['name']}' ({_format_cz_date(active_turnus['end_date'])}).",
                None,
            )

    # B) Zkouší zapsat před začátkem aktivního turnusu
    if active_turnus:
        act_start_dt = dt.datetime.strptime(active_turnus["start_date"], "%Y-%m-%d").date()
        if entry_date < act_start_dt:
            return (
                False,
                f"Záznam pro {formatted_entry} nelze uložit: Datum předchází začátku aktivního turnusu "
                f"'{active_turnus['name']}' ({_format_cz_date(active_turnus['start_date'])}).",
                None,
            )

    # C) Žádný aktivní turnus neběží
    # Najdeme nejnovější uzavřený turnus
    closed_turnusy = [t for t in turnusy if t.get("status") == "closed" and t.get("end_date")]
    if closed_turnusy:
        last_closed = closed_turnusy[0]  # jsou seřazeny sestupně
        last_end_dt = dt.datetime.strptime(last_closed["end_date"], "%Y-%m-%d").date()
        if entry_date > last_end_dt:
            return (
                False,
                f"Záznam pro {formatted_entry} nelze uložit: Turnus '{last_closed['name']}' byl ukončen "
                f"k {_format_cz_date(last_closed['end_date'])}. Po tomto datu nelze zapisovat údaje, "
                "protože není aktivní žádný nový turnus.",
                None,
            )

    return (
        False,
        f"Záznam pro {formatted_entry} nelze uložit: Pro toto datum není aktivní žádný turnus ani nespadá "
        "do období žádné předchozí akce.",
        None,
    )
