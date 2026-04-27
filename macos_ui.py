from __future__ import annotations

import subprocess
from pathlib import Path


class UserCancelled(Exception):
    pass


def _escape_applescript(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def run_applescript(script: str) -> str:
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            check=True,
        )
    except FileNotFoundError as exc:
        raise RuntimeError("macOS dialogs are available only on macOS with osascript.") from exc
    except subprocess.CalledProcessError as exc:
        stderr = (exc.stderr or "").strip()
        if "User canceled" in stderr or "(-128)" in stderr:
            raise UserCancelled() from exc
        raise RuntimeError(stderr or "Не удалось выполнить macOS-диалог.") from exc

    return result.stdout.strip()


def choose_mode() -> bool:
    script = """
    set selectedAction to button returned of (display dialog "Что нужно сделать?" buttons {"Расшифровать", "Обезличить"} default button "Обезличить" with icon note)
    selectedAction
    """
    return run_applescript(script) == "Расшифровать"


def choose_excel_file(prompt: str) -> Path:
    script = f'''
    set selectedFile to choose file with prompt "{_escape_applescript(prompt)}" of type {{"org.openxmlformats.spreadsheetml.sheet","org.openxmlformats.spreadsheetml.template","com.microsoft.excel.xlsm","com.microsoft.excel.xltm"}}
    POSIX path of selectedFile
    '''
    return Path(run_applescript(script))


def show_info(title: str, message: str) -> None:
    script = f'''
    display dialog "{_escape_applescript(message)}" with title "{_escape_applescript(title)}" buttons {{"OK"}} default button "OK" with icon note
    '''
    run_applescript(script)


def show_error(title: str, message: str) -> None:
    script = f'''
    display dialog "{_escape_applescript(message)}" with title "{_escape_applescript(title)}" buttons {{"OK"}} default button "OK" with icon stop
    '''
    run_applescript(script)


def choose_post_action(title: str, message: str, output_path: Path) -> None:
    folder_path = output_path.parent
    output_path_str = _escape_applescript(str(output_path))
    folder_path_str = _escape_applescript(str(folder_path))
    script = f'''
    set selectedAction to button returned of (display dialog "{_escape_applescript(message)}" with title "{_escape_applescript(title)}" buttons {{"Готово", "Открыть папку", "Открыть файл"}} default button "Открыть папку" with icon note)
    if selectedAction is "Открыть файл" then
        do shell script "open " & quoted form of "{output_path_str}"
    else if selectedAction is "Открыть папку" then
        do shell script "open " & quoted form of "{folder_path_str}"
    end if
    '''
    run_applescript(script)
