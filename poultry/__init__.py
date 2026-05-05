from __future__ import annotations

import os
from flask import Flask

from .extensions import db
from .models import (
    Batch,
    DailyRecord,
    VaccineEvent,
    ManpowerLog,
    UtilityReading,
    HarvestEvent,
)


def create_app(database_uri: str | None = None) -> Flask:
    app = Flask(__name__, template_folder="templates", static_folder="static")
    app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret-change-me")
    app.config["SQLALCHEMY_DATABASE_URI"] = (
        database_uri
        or os.environ.get("DATABASE_URL")
        or "sqlite:///" + os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "poultry.db"))
    )
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

    db.init_app(app)

    from .routes import register_routes

    register_routes(app)

    with app.app_context():
        db.create_all()

    return app
