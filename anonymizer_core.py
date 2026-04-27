from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

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
MAPPING_SHEET_NAME = "Расшифровка"


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
        current_index = self._name_index
        self._name_index += 1

        first_name = FIRST_NAMES[current_index % len(FIRST_NAMES)]
        last_name = LAST_NAMES[(current_index // len(FIRST_NAMES)) % len(LAST_NAMES)]

        if parts_count == 2:
            cycle = current_index // (len(FIRST_NAMES) * len(LAST_NAMES))
            candidate = f"{first_name} {last_name}"
            if cycle:
                candidate = f"{candidate} {cycle}"
        else:
            middle_name = MIDDLE_NAMES[
                (current_index // (len(FIRST_NAMES) * len(LAST_NAMES))) % len(MIDDLE_NAMES)
            ]
            cycle = current_index // (len(FIRST_NAMES) * len(LAST_NAMES) * len(MIDDLE_NAMES))
            candidate = f"{last_name} {first_name} {middle_name}"
            if cycle:
                candidate = f"{candidate} {cycle}"

        self._used_names.add(candidate)
        return candidate

    def _generate_fake_login(self) -> str:
        self._login_index += 1
        return f"user{self._login_index:04d}.masked"


def remove_hyperlink(cell) -> None:
    if cell.hyperlink is not None:
        cell.hyperlink = None


def iter_existing_cells(worksheet):
    cells = getattr(worksheet, "_cells", None)
    if isinstance(cells, dict) and cells:
        return cells.values()

    return (
        cell
        for row in worksheet.iter_rows()
        for cell in row
        if cell.value is not None or cell.hyperlink is not None
    )


def validate_excel_path(input_path: Path) -> Path:
    resolved_path = input_path.expanduser().resolve()
    if resolved_path.suffix.lower() not in SUPPORTED_EXCEL_SUFFIXES:
        raise ValueError("Поддерживаются только Excel-файлы: .xlsx, .xlsm, .xltx, .xltm")
    if not resolved_path.exists():
        raise FileNotFoundError(f"Файл не найден: {resolved_path}")
    return resolved_path


def anonymize_workbook(input_path: Path, output_path: Path) -> tuple[Path, Anonymizer]:
    workbook = load_workbook(input_path)
    anonymizer = Anonymizer()

    for worksheet in workbook.worksheets:
        for cell in iter_existing_cells(worksheet):
            remove_hyperlink(cell)
            if isinstance(cell.value, str):
                cell.value = anonymizer.anonymize_text(cell.value)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return output_path, anonymizer


def save_mapping_workbook(
    output_path: Path, anonymizer: Anonymizer, sheet_name: str = MAPPING_SHEET_NAME
) -> Path:
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


def load_reverse_mapping(mapping_path: Path) -> list[tuple[str, str]]:
    workbook = load_workbook(mapping_path, read_only=True, data_only=True)
    worksheet = workbook[MAPPING_SHEET_NAME] if MAPPING_SHEET_NAME in workbook.sheetnames else workbook.active
    replacements: list[tuple[str, str]] = []

    for row in worksheet.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 3:
            continue

        _, original_value, masked_value = row[:3]
        if not isinstance(original_value, str) or not isinstance(masked_value, str):
            continue
        if not masked_value:
            continue
        replacements.append((masked_value, original_value))

    workbook.close()
    return sorted(replacements, key=lambda item: len(item[0]), reverse=True)


def deanonymize_text(value: str, replacements: Iterable[tuple[str, str]]) -> str:
    restored_value = value
    for masked_value, original_value in replacements:
        restored_value = restored_value.replace(masked_value, original_value)
    return restored_value


def deanonymize_workbook(input_path: Path, mapping_path: Path, output_path: Path) -> Path:
    workbook = load_workbook(input_path)
    replacements = load_reverse_mapping(mapping_path)

    for worksheet in workbook.worksheets:
        for cell in iter_existing_cells(worksheet):
            if isinstance(cell.value, str):
                cell.value = deanonymize_text(cell.value, replacements)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return output_path


def build_output_dir(input_path: Path, suffix: str) -> Path:
    return input_path.parent / f"{input_path.stem}_{suffix}"


def build_output_paths(input_path: Path) -> tuple[Path, Path]:
    output_dir = build_output_dir(input_path, "anonymized")
    anonymized_path = output_dir / f"{input_path.stem}_anonymized{input_path.suffix}"
    mapping_path = output_dir / f"{input_path.stem}_mapping.xlsx"
    return anonymized_path, mapping_path


def build_decrypted_output_path(input_path: Path) -> Path:
    output_dir = build_output_dir(input_path, "decrypted")
    return output_dir / f"{input_path.stem}_decrypted{input_path.suffix}"


def infer_mapping_path(input_path: Path) -> Path:
    stem = input_path.stem
    original_stem = stem.removesuffix("_anonymized")
    candidate = input_path.parent / f"{original_stem}_mapping.xlsx"
    if candidate.exists():
        return candidate
    raise FileNotFoundError(
        "Не удалось автоматически найти файл расшифровки. "
        "Передайте путь через аргумент --mapping-file."
    )
