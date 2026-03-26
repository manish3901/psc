from psc_coupens_app import create_app, db


def main():
    app = create_app()
    with app.app_context():
        db.create_all()
        print("Database tables created/verified.")


if __name__ == "__main__":
    main()

