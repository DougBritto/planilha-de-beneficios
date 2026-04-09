from consolidador import create_app

app = create_app()


if __name__ == "__main__":
    settings = app.config["APP_SETTINGS"]
    app.run(host="0.0.0.0", port=settings.port, debug=settings.debug)
