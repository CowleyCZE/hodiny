"""Webové routy autentizace a správy uživatelů."""

from flask import Blueprint, flash, g, redirect, render_template, request, url_for

from services.auth_service import login_user, logout_user, role_required


auth_bp = Blueprint("auth", __name__, url_prefix="/auth")


@auth_bp.route("/setup", methods=["GET", "POST"])
def setup():
    if g.database.user_count() > 0:
        return redirect(url_for("auth.login"))
    if request.method == "POST":
        try:
            g.database.create_user(request.form.get("username", ""), request.form.get("password", ""), "admin")
            flash("Administrátorský účet byl vytvořen. Nyní se přihlaste.", "success")
            return redirect(url_for("auth.login"))
        except ValueError as error:
            flash(str(error), "error")
    return render_template("auth_setup.html")


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = g.database.authenticate(request.form.get("username", ""), request.form.get("password", ""))
        if user:
            login_user(user)
            return redirect(request.args.get("next") or url_for("main.index"))
        flash("Neplatné přihlašovací údaje.", "error")
    return render_template("login.html")


@auth_bp.post("/logout")
def logout():
    logout_user()
    return redirect(url_for("auth.login"))


@auth_bp.route("/users", methods=["GET", "POST"])
@role_required("admin")
def users():
    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "create":
                g.database.create_user(
                    request.form.get("username", ""),
                    request.form.get("password", ""),
                    request.form.get("role", "viewer"),
                )
                flash("Uživatel byl vytvořen.", "success")
            elif action in {"activate", "deactivate"}:
                g.database.set_user_active(int(request.form["user_id"]), action == "activate")
                flash("Stav uživatele byl změněn.", "success")
        except (ValueError, KeyError) as error:
            flash(str(error), "error")
    return render_template("users.html", users=g.database.list_users())
