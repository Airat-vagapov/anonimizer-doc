from __future__ import annotations

import subprocess
import sys
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except ImportError as exc:
    raise ImportError(
        "tkinter не установлен. "
        "Linux (Debian/Ubuntu): sudo apt install python3-tk"
    ) from exc


class UserCancelled(Exception):
    pass


class _ButtonDialog(tk.Toplevel):
    """Modal dialog with arbitrary buttons. result is the label of the clicked button or None."""

    def __init__(self, parent: tk.Tk, title: str, message: str, buttons: list[str], default: str):
        super().__init__(parent)
        self.result: str | None = None
        self.title(title)
        self.resizable(False, False)

        tk.Label(self, text=message, wraplength=360, justify="left", pady=14, padx=20).pack()

        frame = tk.Frame(self)
        frame.pack(pady=(0, 14), padx=20)

        for label in buttons:
            tk.Button(
                frame,
                text=label,
                width=15,
                command=lambda lbl=label: self._pick(lbl),
                default="active" if label == default else "normal",
            ).pack(side=tk.LEFT, padx=4)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.grab_set()
        self.focus_set()
        self.lift()

    def _pick(self, label: str) -> None:
        self.result = label
        self.destroy()


def _dialog(title: str, message: str, buttons: list[str], default: str) -> str | None:
    root = tk.Tk()
    root.withdraw()
    dlg = _ButtonDialog(root, title, message, buttons, default)
    root.wait_window(dlg)
    result = dlg.result
    root.destroy()
    return result


def choose_mode() -> bool:
    result = _dialog(
        "Excel Anonymizer",
        "Что нужно сделать?",
        ["Обезличить", "Расшифровать"],
        "Обезличить",
    )
    if result is None:
        raise UserCancelled()
    return result == "Расшифровать"


def choose_excel_file(prompt: str) -> Path:
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title=prompt,
        filetypes=[
            ("Excel files", "*.xlsx *.xlsm *.xltx *.xltm"),
            ("All files", "*.*"),
        ],
    )
    root.destroy()
    if not path:
        raise UserCancelled()
    return Path(path)


def show_error(title: str, message: str) -> None:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(title, message)
    root.destroy()


def choose_post_action(title: str, message: str, output_path: Path) -> None:
    result = _dialog(title, message, ["Готово", "Открыть папку", "Открыть файл"], "Открыть папку")
    if result == "Открыть файл":
        _open(output_path)
    elif result == "Открыть папку":
        _open(output_path.parent)


def _open(path: Path) -> None:
    if sys.platform == "win32":
        subprocess.run(["explorer", str(path)], check=False)
    else:
        subprocess.run(["xdg-open", str(path)], check=False)
