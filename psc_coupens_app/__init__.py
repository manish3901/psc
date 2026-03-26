import os

from dotenv import load_dotenv
from flask import Flask
from flask_sqlalchemy import SQLAlchemy

load_dotenv()

db = SQLAlchemy()


def create_app():
    app = Flask(__name__)

    app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "dev_secret")
    db_url = os.getenv("DATABASE_URL", "sqlite:///psc_coupens.db")
    # Some providers still supply "postgres://" URLs; SQLAlchemy expects "postgresql://".
    if db_url.startswith("postgres://"):
        db_url = "postgresql://" + db_url[len("postgres://") :]
    app.config["SQLALCHEMY_DATABASE_URI"] = db_url
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

    # SQLite tuning for small deployments (reduces "database is locked" errors).
    if db_url.startswith("sqlite:"):
        app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {"connect_args": {"timeout": 30}}

    db.init_app(app)

    from .routes import main

    app.register_blueprint(main)
    return app
