import pytest

import app as app_module
from config import Config
from services.database import Database


@pytest.fixture
def secured_client(tmp_path, monkeypatch):
    database = Database(tmp_path / "hodiny.sqlite3")
    monkeypatch.setattr(app_module, "database", database)
    monkeypatch.setattr(Config, "AUTH_REQUIRED", True)
    app_module.app.config.update(TESTING=True, WTF_CSRF_ENABLED=False)
    with app_module.app.test_client() as client:
        yield client, database


def test_database_hashes_password_and_supports_roles(tmp_path):
    database = Database(tmp_path / "db.sqlite3")
    database.create_user("Admin", "tajneheslo", "admin")
    assert database.authenticate("admin", "tajneheslo")["role"] == "admin"
    assert database.authenticate("admin", "spatne") is None
    with database.connection() as connection:
        password_hash = connection.execute("SELECT password_hash FROM users").fetchone()[0]
    assert "tajneheslo" not in password_hash


def test_first_setup_login_and_protected_sync_api(secured_client):
    client, database = secured_client
    response = client.get("/api/v1/sync/status")
    assert response.status_code == 302
    assert client.get("/auth/setup").status_code == 200
    response = client.post("/auth/setup", data={"username": "admin", "password": "tajneheslo"})
    assert response.status_code == 302
    assert database.user_count() == 1
    client.post("/auth/login", data={"username": "admin", "password": "tajneheslo"})
    response = client.get("/api/v1/sync/status")
    assert response.status_code == 200
    assert response.get_json()["success"] is True


def test_health_is_public_and_sets_security_headers(secured_client):
    client, _ = secured_client
    response = client.get("/api/v1/health")
    assert response.status_code == 200
    assert response.headers["X-Content-Type-Options"] == "nosniff"
    assert response.headers["Permissions-Policy"]
