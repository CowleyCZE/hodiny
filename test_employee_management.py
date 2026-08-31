import json

from employee_management import EmployeeManager


def test_preferred_employee_is_first_and_auto_selected(tmp_path):
    data_path = tmp_path / "data"
    data_path.mkdir(parents=True, exist_ok=True)
    (data_path / "employee_config.json").write_text(
        json.dumps(
            {
                "zamestnanci": ["Beta Worker", "Jan Test", "Alpha Worker"],
                "vybrani_zamestnanci": ["Alpha Worker"],
            }
        ),
        encoding="utf-8",
    )

    manager = EmployeeManager(data_path, preferred_employee_name="Jan Test")

    assert [employee["name"] for employee in manager.get_all_employees()] == ["Jan Test", "Alpha Worker", "Beta Worker"]
    assert manager.get_vybrani_zamestnanci() == ["Jan Test", "Alpha Worker"]


def test_preferred_employee_cannot_be_removed_from_selected_list(tmp_path):
    data_path = tmp_path / "data"
    data_path.mkdir(parents=True, exist_ok=True)
    (data_path / "employee_config.json").write_text(
        json.dumps(
            {
                "zamestnanci": ["Jan Test", "Alpha Worker"],
                "vybrani_zamestnanci": ["Jan Test", "Alpha Worker"],
            }
        ),
        encoding="utf-8",
    )

    manager = EmployeeManager(data_path, preferred_employee_name="Jan Test")

    assert manager.odebrat_vybraneho_zamestnance("Jan Test") is False
    assert manager.get_vybrani_zamestnanci()[0] == "Jan Test"


def test_corrupt_json_creates_backup(tmp_path):
    data_path = tmp_path / "data"
    data_path.mkdir(parents=True, exist_ok=True)
    config_file = data_path / "employee_config.json"
    config_file.write_text("{ invalid json ...", encoding="utf-8")

    manager = EmployeeManager(data_path, preferred_employee_name="Jan Test")

    assert manager.get_employee_names() == []
    backup_file = data_path / "employee_config.json.bak"
    assert backup_file.exists()
    assert backup_file.read_text(encoding="utf-8") == "{ invalid json ..."


def test_rename_preferred_employee_updates_name(tmp_path):
    data_path = tmp_path / "data"
    data_path.mkdir(parents=True, exist_ok=True)
    (data_path / "employee_config.json").write_text(
        json.dumps(
            {
                "zamestnanci": ["Jan Stary", "Petr Pospisil"],
                "vybrani_zamestnanci": ["Jan Stary"],
            }
        ),
        encoding="utf-8",
    )

    manager = EmployeeManager(data_path, preferred_employee_name="Jan Stary")
    assert manager.upravit_zamestnance_podle_jmena("Jan Stary", "Jan Novy") is True
    assert manager.preferred_employee_name == "Jan Novy"
    assert manager.get_vybrani_zamestnanci()[0] == "Jan Novy"
