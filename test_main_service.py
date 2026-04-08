from unittest.mock import Mock

import pytest

from services.main_service import save_time_entry


def test_save_time_entry_raises_when_excel_write_fails():
    excel_manager = Mock()
    excel_manager.ulozit_pracovni_dobu.return_value = False
    hodiny2025_manager = Mock()

    with pytest.raises(OSError, match="Nepodařilo se uložit pracovní dobu do Excel souboru"):
        save_time_entry(
            excel_manager,
            hodiny2025_manager,
            "2026-04-09",
            "08:00",
            "16:00",
            "0.5",
            ["Test Zamestnanec"],
            False,
        )

    hodiny2025_manager.zapis_pracovni_doby.assert_not_called()


def test_save_time_entry_does_not_double_write_hodiny2025_on_success():
    excel_manager = Mock()
    excel_manager.ulozit_pracovni_dobu.return_value = True
    hodiny2025_manager = Mock()

    message = save_time_entry(
        excel_manager,
        hodiny2025_manager,
        "2026-04-09",
        "08:00",
        "16:00",
        "0.5",
        ["Test Zamestnanec"],
        False,
    )

    assert "Pracovní doba" in message
    excel_manager.ulozit_pracovni_dobu.assert_called_once_with(
        "2026-04-09",
        "08:00",
        "16:00",
        "0.5",
        ["Test Zamestnanec"],
    )
    hodiny2025_manager.zapis_pracovni_doby.assert_not_called()


def test_save_time_entry_raises_when_free_day_excel_write_fails():
    excel_manager = Mock()
    excel_manager.ulozit_pracovni_dobu.return_value = False
    hodiny2025_manager = Mock()

    with pytest.raises(OSError, match="Nepodařilo se uložit volný den do Excel souboru"):
        save_time_entry(
            excel_manager,
            hodiny2025_manager,
            "2026-04-09",
            None,
            None,
            "0",
            ["Test Zamestnanec"],
            True,
        )

    hodiny2025_manager.zapis_pracovni_doby.assert_not_called()
