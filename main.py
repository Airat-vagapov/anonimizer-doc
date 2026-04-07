from __future__ import annotations

import argparse
from pathlib import Path

from anonymizer_core import (
    anonymize_workbook,
    build_decrypted_output_path,
    build_output_paths,
    deanonymize_workbook,
    infer_mapping_path,
    save_mapping_workbook,
    validate_excel_path,
)
from macos_ui import UserCancelled, choose_excel_file, choose_mode, choose_post_action, show_error


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Обезличивание и расшифровка Excel-файла")
    parser.add_argument("input_file", nargs="?", help="Путь до Excel-файла")
    parser.add_argument(
        "--decrypt",
        action="store_true",
        help="Расшифровать ранее обезличенный Excel-файл",
    )
    parser.add_argument(
        "--mapping-file",
        help="Путь до файла соответствий *_mapping.xlsx для режима --decrypt",
    )
    return parser.parse_args()


def run_anonymize(input_path: Path) -> tuple[Path, Path]:
    output_path, mapping_path = build_output_paths(input_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    anonymized_file, anonymizer = anonymize_workbook(input_path, output_path)
    mapping_file = save_mapping_workbook(mapping_path, anonymizer)
    return anonymized_file, mapping_file


def run_decrypt(input_path: Path, mapping_path: Path) -> Path:
    output_path = build_decrypted_output_path(input_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    return deanonymize_workbook(input_path, mapping_path, output_path)


def cli_main(args: argparse.Namespace) -> int:
    input_path = validate_excel_path(Path(args.input_file)) if args.input_file else None

    if args.decrypt:
        if input_path is None:
            raise ValueError("Для режима --decrypt нужно передать путь до Excel-файла.")
        mapping_path = validate_excel_path(Path(args.mapping_file)) if args.mapping_file else infer_mapping_path(input_path)
        decrypted_file = run_decrypt(input_path, mapping_path)
        print(f"Расшифрованный файл сохранен: {decrypted_file}")
        print(f"Использован файл расшифровки: {mapping_path}")
        return 0

    if input_path is None:
        raise ValueError("Передайте путь до Excel-файла или запустите без аргументов для оконного режима.")

    anonymized_file, mapping_file = run_anonymize(input_path)
    print(f"Обезличенный файл сохранен: {anonymized_file}")
    print(f"Файл расшифровки сохранен: {mapping_file}")
    return 0


def gui_main() -> int:
    decrypt_mode = choose_mode()
    input_prompt = "Выберите Excel-файл для расшифровки" if decrypt_mode else "Выберите Excel-файл для обезличивания"
    input_path = validate_excel_path(choose_excel_file(input_prompt))

    if decrypt_mode:
        try:
            mapping_path = infer_mapping_path(input_path)
        except FileNotFoundError:
            mapping_path = validate_excel_path(choose_excel_file("Выберите файл расшифровки mapping.xlsx"))

        decrypted_file = run_decrypt(input_path, mapping_path)
        choose_post_action(
            "Расшифровка завершена",
            f"Готово.\n\nФайл:\n{decrypted_file}\n\nMapping:\n{mapping_path}",
            decrypted_file,
        )
        return 0

    anonymized_file, mapping_file = run_anonymize(input_path)
    choose_post_action(
        "Обезличивание завершено",
        f"Готово.\n\nОбезличенный файл:\n{anonymized_file}\n\nФайл соответствий:\n{mapping_file}",
        anonymized_file,
    )
    return 0


def main() -> int:
    args = parse_args()

    try:
        if args.input_file or args.decrypt or args.mapping_file:
            return cli_main(args)
        return gui_main()
    except UserCancelled:
        return 1
    except Exception as exc:
        if args.input_file or args.decrypt or args.mapping_file:
            print(f"Ошибка: {exc}")
        else:
            try:
                show_error("Ошибка", str(exc))
            except Exception:
                print(f"Ошибка: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
