from __future__ import annotations

import os
import sys
import time
import threading
import webbrowser
import urllib.request
from pathlib import Path


APP_PORT = 8501
APP_HOST = "127.0.0.1"
APP_URL = f"http://{APP_HOST}:{APP_PORT}"


def get_base_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)

    return Path(__file__).resolve().parent


def wait_for_streamlit_and_open_browser():
    for _ in range(120):
        try:
            with urllib.request.urlopen(APP_URL, timeout=1) as response:
                if response.status in (200, 302):
                    webbrowser.open_new(APP_URL)
                    return
        except Exception:
            time.sleep(0.5)

    print(f"Streamlit UI wurde nicht erreichbar unter {APP_URL}")


def main() -> int:
    base_path = get_base_path()
    app_path = base_path / "front.py"

    if not app_path.exists():
        print(f"front.py wurde nicht gefunden: {app_path}")
        input("Enter drücken zum Schließen...")
        return 1

    try:
        os.chdir(base_path)

        for key in list(os.environ.keys()):
            if key.startswith("STREAMLIT_"):
                os.environ.pop(key, None)

        os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
        os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
        os.environ["STREAMLIT_SERVER_ADDRESS"] = APP_HOST
        os.environ["STREAMLIT_SERVER_PORT"] = str(APP_PORT)
        os.environ["STREAMLIT_SERVER_BASE_URL_PATH"] = ""

        os.environ["STREAMLIT_BROWSER_SERVER_ADDRESS"] = APP_HOST
        os.environ["STREAMLIT_BROWSER_SERVER_PORT"] = str(APP_PORT)
        os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

        os.environ["STREAMLIT_SERVER_ENABLE_CORS"] = "true"
        os.environ["STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION"] = "false"
        os.environ["STREAMLIT_LOGGER_LEVEL"] = "info"

        import streamlit.config as config
        from streamlit.web import bootstrap

        config.set_option("global.developmentMode", False)

        config.set_option("server.headless", True)
        config.set_option("server.address", APP_HOST)
        config.set_option("server.port", APP_PORT)
        config.set_option("server.baseUrlPath", "")
        config.set_option("server.enableCORS", True)
        config.set_option("server.enableXsrfProtection", False)

        config.set_option("browser.serverAddress", APP_HOST)
        config.set_option("browser.serverPort", APP_PORT)
        config.set_option("browser.gatherUsageStats", False)

        config.set_option("logger.level", "info")

        threading.Thread(
            target=wait_for_streamlit_and_open_browser,
            daemon=True,
        ).start()

        flag_options = {
            "global.developmentMode": False,

            "server.headless": True,
            "server.address": APP_HOST,
            "server.port": APP_PORT,
            "server.baseUrlPath": "",
            "server.enableCORS": True,
            "server.enableXsrfProtection": False,

            "browser.serverAddress": APP_HOST,
            "browser.serverPort": APP_PORT,
            "browser.gatherUsageStats": False,

            "logger.level": "info",
        }

        bootstrap.run(
            str(app_path),
            False,
            [],
            flag_options,
        )

        return 0

    except Exception as error:
        print("Fehler beim Starten der App:")
        print(error)
        input("Enter drücken zum Schließen...")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())