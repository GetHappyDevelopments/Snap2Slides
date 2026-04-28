from __future__ import annotations

import importlib.util
import os
import subprocess
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk


APP_ROOT = Path(__file__).resolve().parent
SPLASH_PATH = APP_ROOT / "assets" / "splash.png"
REQUIREMENTS_PATH = APP_ROOT / "requirements.txt"
MIN_SPLASH_SECONDS = 3.0

REQUIRED_MODULES = {
    "pillow": "PIL",
    "python-pptx": "pptx",
    "numpy": "numpy",
    "opencv-python": "cv2",
    "rapidocr-onnxruntime": "rapidocr_onnxruntime",
}


class SplashWindow(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.overrideredirect(True)
        self.configure(background="#020b1c")
        self._start_time = time.monotonic()
        self._image = None

        self._build_ui()
        self.after(50, self._start_worker)

    def _build_ui(self) -> None:
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        image_label: tk.Label | None = None
        if SPLASH_PATH.exists():
            image = tk.PhotoImage(file=str(SPLASH_PATH))
            factor = max(1, int(max(image.width() / min(980, screen_w * 0.82), image.height() / min(560, screen_h * 0.72))) + 1)
            if factor > 1:
                image = image.subsample(factor, factor)
            self._image = image
            image_label = tk.Label(self, image=image, borderwidth=0, highlightthickness=0)
            image_label.pack()
        else:
            fallback = ttk.Frame(self, padding=32)
            fallback.pack()
            ttk.Label(fallback, text="Snap2Slides", font=("Segoe UI", 28, "bold")).pack()

        status_bar = ttk.Frame(self, padding=(18, 10, 18, 14))
        status_bar.pack(fill=tk.X)
        self.status = tk.StringVar(value="Snap2Slides wird vorbereitet...")
        ttk.Label(status_bar, textvariable=self.status).pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(status_bar, mode="indeterminate", length=180)
        self.progress.pack(side=tk.RIGHT)
        self.progress.start(12)

        self.update_idletasks()
        width = self.winfo_width() if image_label is not None else 420
        height = self.winfo_height()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.geometry(f"+{x}+{y}")

    def _start_worker(self) -> None:
        threading.Thread(target=self._prepare_and_start, daemon=True).start()

    def _prepare_and_start(self) -> None:
        try:
            self._install_missing_dependencies()
        except Exception as exc:
            self.after(0, self._show_error, exc)
            return

        remaining_ms = max(0, int((MIN_SPLASH_SECONDS - (time.monotonic() - self._start_time)) * 1000))
        self.after(remaining_ms, self._launch_app)

    def _install_missing_dependencies(self) -> None:
        missing = [package for package, module in REQUIRED_MODULES.items() if importlib.util.find_spec(module) is None]
        if not missing:
            self._set_status("Abhaengigkeiten bereit. Starte Programm...")
            return

        self._set_status("Installiere Abhaengigkeiten...")
        command = [sys.executable, "-m", "pip", "install", "--user", "-r", str(REQUIREMENTS_PATH)]
        result = subprocess.run(command, cwd=APP_ROOT, text=True, capture_output=True)
        if result.returncode != 0:
            details = result.stderr.strip() or result.stdout.strip() or "Unbekannter pip-Fehler."
            raise RuntimeError(details)
        self._set_status("Installation abgeschlossen. Starte Programm...")

    def _set_status(self, message: str) -> None:
        self.after(0, self.status.set, message)

    def _show_error(self, exc: Exception) -> None:
        self.progress.stop()
        messagebox.showerror("Snap2Slides konnte nicht gestartet werden", str(exc))
        self.destroy()

    def _launch_app(self) -> None:
        self.progress.stop()
        self.destroy()
        os.environ["SNAP2SLIDES_SKIP_SPLASH"] = "1"
        from app import main

        main()


def main() -> None:
    SplashWindow().mainloop()


if __name__ == "__main__":
    main()
