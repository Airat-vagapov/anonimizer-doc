from __future__ import annotations

import argparse
import re
import subprocess
from pathlib import Path

from openpyxl import Workbook, load_workbook

FIRST_NAMES = [
    "Алексей",
    "Андрей",
    "Анна",
    "Виктор",
    "Дарья",
    "Екатерина",
    "Ирина",
    "Мария",
    "Никита",
    "Ольга",
    "Сергей",
    "Татьяна",
]

LAST_NAMES = [
    "Александров",
    "Белов",
    "Волков",
    "Громов",
    "Егоров",
    "Захаров",
    "Ильин",
    "Крылов",
    "Морозов",
    "Назаров",
    "Орлов",
    "Соколов",
]

MIDDLE_NAMES = [
    "Алексеевич",
    "Андреевич",
    "Викторович",
    "Дмитриевич",
    "Игоревич",
    "Олегович",
    "Сергеевич",
    "Алексеевна",
    "Андреевна",
    "Викторовна",
    "Дмитриевна",
    "Игоревна",
    "Олеговна",
    "Сергеевна",
]

FIO_PATTERN = re.compile(
    r"\b[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?(?:\s+[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?){1,2}\b"
)
JIRA_LOGIN_PATTERN = re.compile(r"\b[a-z][a-z0-9_-]*\.[a-z][a-z0-9_-]*\b", re.IGNORECASE)
SUPPORTED_EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}


class Anonymizer:
    def __init__(self) -> None:
        self.full_name_map: dict[str, str] = {}
        self.login_map: dict[str, str] = {}
        self._used_names: set[str] = set()
        self._name_index = 0
        self._login_index = 0

    def anonymize_text(self, value: str) -> str:
        value = self._replace_full_names(value)
        value = self._replace_logins(value)
        return value

    def _replace_full_names(self, value: str) -> str:
        return FIO_PATTERN.sub(self._full_name_replacer, value)

    def _replace_logins(self, value: str) -> str:
        return JIRA_LOGIN_PATTERN.sub(self._login_replacer, value)

    def _full_name_replacer(self, match: re.Match[str]) -> str:
        original_name = match.group(0)
        return self.full_name_map.setdefault(original_name, self._generate_fake_name(original_name))

    def _login_replacer(self, match: re.Match[str]) -> str:
        original_login = match.group(0)
        return self.login_map.setdefault(original_login, self._generate_fake_login())

    def _generate_fake_name(self, original_name: str) -> str:
        parts_count = len(original_name.split())

        while True:
            first_name = FIRST_NAMES[self._name_index % len(FIRST_NAMES)]
            last_name = LAST_NAMES[(self._name_index // len(FIRST_NAMES)) % len(LAST_NAMES)]
            middle_name = MIDDLE_NAMES[
                (self._name_index // (len(FIRST_NAMES) * len(LAST_NAMES))) % len(MIDDLE_NAMES)
            ]
            self._name_index += 1

            if parts_count == 2:
                candidate = f"{first_name} {last_name}"
            else:
                candidate = f"{last_name} {first_name} {middle_name}"

            if candidate not in self._used_names:
                self._used_names.add(candidate)
                return candidate

    def _generate_fake_login(self) -> str:
        self._login_index += 1
        return f"user{self._login_index:04d}.masked"


def remove_hyperlink(cell) -> None:
    if cell.hyperlink is not None:
        cell.hyperlink = None


def anonymize_workbook(input_path: Path, output_path: Path) -> tuple[Path, Anonymizer]:
    workbook = load_workbook(input_path)
    anonymizer = Anonymizer()

    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                remove_hyperlink(cell)
                if isinstance(cell.value, str):
                    cell.value = anonymizer.anonymize_text(cell.value)

    workbook.save(output_path)
    return output_path, anonymizer


def save_mapping_workbook(output_path: Path, anonymizer: Anonymizer, sheet_name: str = "Расшифровка") -> Path:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    worksheet.append(["Тип", "Оригинал", "Замена"])

    for original_name, fake_name in anonymizer.full_name_map.items():
        worksheet.append(["ФИО", original_name, fake_name])

    for original_login, fake_login in anonymizer.login_map.items():
        worksheet.append(["Логин", original_login, fake_login])

    for column in ("A", "B", "C"):
        worksheet.column_dimensions[column].width = 32

    workbook.save(output_path)
    return output_path


def build_output_dir(input_path: Path) -> Path:
    return input_path.parent / f"{input_path.stem}_anonymized"


def build_output_paths(input_path: Path) -> tuple[Path, Path]:
    output_dir = build_output_dir(input_path)
    anonymized_path = output_dir / f"{input_path.stem}_anonymized{input_path.suffix}"
    mapping_path = output_dir / f"{input_path.stem}_mapping.xlsx"
    return anonymized_path, mapping_path


def choose_input_file() -> Path:
    script = """
    set selectedFile to choose file with prompt "Выберите Excel-файл для обезличивания"
    POSIX path of selectedFile
    """
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            check=True,
        )
        selected_path = result.stdout.strip()
        if selected_path:
            input_path = Path(selected_path).expanduser().resolve()
            if input_path.suffix.lower() not in SUPPORTED_EXCEL_SUFFIXES:
                raise ValueError(
                    "Поддерживаются только Excel-файлы: .xlsx, .xlsm, .xltx, .xltm"
                )
            return input_path
    except (subprocess.CalledProcessError, FileNotFoundError):
        pass

    manual_path = input("Введите путь до Excel-файла: ").strip()
    if not manual_path:
        raise ValueError("Не выбран входной файл")
    input_path = Path(manual_path).expanduser().resolve()
    if input_path.suffix.lower() not in SUPPORTED_EXCEL_SUFFIXES:
        raise ValueError("Поддерживаются только Excel-файлы: .xlsx, .xlsm, .xltx, .xltm")
    return input_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Обезличивание Excel-файла")
    parser.add_argument("input_file", nargs="?", help="Путь до исходного Excel-файла")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = (
        Path(args.input_file).expanduser().resolve()
        if args.input_file
        else choose_input_file()
    )
    output_path, mapping_path = build_output_paths(input_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    anonymized_file, anonymizer = anonymize_workbook(input_path, output_path)
    mapping_file = save_mapping_workbook(mapping_path, anonymizer)

    print(f"Обезличенный файл сохранен: {anonymized_file}")
    print(f"Файл расшифровки сохранен: {mapping_file}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
