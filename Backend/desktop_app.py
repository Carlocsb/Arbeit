import json
import os
import socket
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path
import uvicorn
import webview


def find_free_port(start_port: int = 8000, end_port: int = 8100) -> int:
    """
    Sucht einen freien lokalen Port.
    Falls 8000 belegt ist, wird automatisch 8001, 8002, ... verwendet.
    """
    for port in range(start_port, end_port + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.2)
            if sock.connect_ex(("127.0.0.1", port)) != 0:
                return port

    raise RuntimeError("Kein freier Port zwischen 8000 und 8100 gefunden.")


def configure_environment(host: str, port: int) -> str:
    """
    Setzt die Backend-URL, bevor main.py importiert wird.

    Wichtig:
    main.py liest PPT_AUTOMATION_BASE_URL beim Import ein.
    Deshalb darf `from main import app` nicht oben auf Modulebene stehen.
    """
    base_url = f"http://{host}:{port}"
    os.environ["PPT_AUTOMATION_HOST"] = host
    os.environ["PPT_AUTOMATION_PORT"] = str(port)
    os.environ["PPT_AUTOMATION_BASE_URL"] = base_url
    return base_url


def start_backend(host: str, port: int):
    """
    Startet das FastAPI-Backend im Hintergrund.

    Die eigentliche Erkennung passiert in main.py:
    - native PowerPoint-Diagramme => python-pptx
    - think-cell-Mappings => PowerPoint/Excel/think-cell COM-Automation
    """
    from main import app

    uvicorn.run(
        app,
        host=host,
        port=port,
        reload=False,
        access_log=False,
    )


def wait_for_backend(base_url: str, timeout_seconds: float = 12.0) -> dict:
    """
    Wartet, bis /health erreichbar ist.
    Gibt die Health-Antwort zurück oder wirft RuntimeError.
    """
    deadline = time.time() + timeout_seconds
    last_error = None

    while time.time() < deadline:
        try:
            with urllib.request.urlopen(f"{base_url}/health", timeout=1.0) as response:
                return json.loads(response.read().decode("utf-8"))
        except Exception as error:
            last_error = error
            time.sleep(0.3)

    raise RuntimeError(f"Backend konnte nicht gestartet werden: {last_error}")


def get_thinkcell_status(base_url: str) -> dict:
    """
    Fragt optional /thinkcell-health ab.
    Wenn die Route in main.py nicht existiert, läuft die App trotzdem weiter.
    """
    try:
        with urllib.request.urlopen(f"{base_url}/thinkcell-health", timeout=3.0) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        return {
            "available": False,
            "reason": f"/thinkcell-health nicht verfügbar: HTTP {error.code}",
        }
    except Exception as error:
        return {
            "available": False,
            "reason": str(error),
        }
class DesktopApi:
    def __init__(self):
        self.window = None

    def save_ppt_file(self, file_url: str, suggested_filename: str = "Aktualisierte_Praesentation.pptx"):
        """
        Öffnet einen nativen Speichern-Dialog und speichert die generierte PPTX
        an dem vom User ausgewählten Ort.
        """
        if self.window is None:
            return {
                "success": False,
                "cancelled": False,
                "message": "Desktop-Fenster ist nicht initialisiert."
            }

        save_path = self.window.create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename=suggested_filename,
            file_types=("PowerPoint (*.pptx)", "Alle Dateien (*.*)")
        )

        if not save_path:
            return {
                "success": False,
                "cancelled": True,
                "message": "Speichern wurde abgebrochen."
            }

        if isinstance(save_path, (list, tuple)):
            save_path = save_path[0]

        try:
            with urllib.request.urlopen(file_url, timeout=60) as response:
                data = response.read()

            output_path = Path(save_path)

            if output_path.suffix.lower() != ".pptx":
                output_path = output_path.with_suffix(".pptx")

            output_path.write_bytes(data)

            return {
                "success": True,
                "cancelled": False,
                "message": f"PowerPoint wurde gespeichert: {output_path}",
                "path": str(output_path)
            }

        except Exception as error:
            return {
                "success": False,
                "cancelled": False,
                "message": str(error)
            }

def main():
    host = "127.0.0.1"
    port = find_free_port()
    base_url = configure_environment(host, port)

    backend_thread = threading.Thread(
        target=start_backend,
        args=(host, port),
        daemon=True,
    )
    backend_thread.start()

    health = wait_for_backend(base_url)
    thinkcell_status = get_thinkcell_status(base_url)

    print("--------------------------------------------------")
    print("PowerPoint Automation Studio startet")
    print(f"URL:             {base_url}")
    print(f"Backend:         {health.get('status', 'unbekannt')}")
    print(f"Frontend-Datei:  {health.get('frontend_file', 'unbekannt')}")
    print(f"think-cell:      {thinkcell_status.get('available')}")
    print(f"think-cell Info: {thinkcell_status.get('reason')}")
    print("--------------------------------------------------")

    webview.create_window(
        title="PowerPoint Automation Studio",
        url=base_url,
        width=1500,
        height=950,
        min_size=(1100, 750),
    )

    webview.start()


if __name__ == "__main__":
    main()
