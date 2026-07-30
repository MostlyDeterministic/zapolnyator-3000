#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import re
import shutil
import sys
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Set, Tuple

from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P
from openpyxl import load_workbook


SHEET_NAME_ODS = "Сигналы"
SHEET_NAME_XLSX = "Печать"
START_ROW = 5

COL_ODS_MODULE_SLOT = "C"
COL_ODS_BASKET = "E"
COL_ODS_SOURCE_FILE = "J"
COL_ODS_CHANNEL = "R"
COL_ODS_SIGNAL_TYPE = "V"

COL_ODS_TARGET_P = "S"
COL_ODS_TARGET_R = "T"
COL_ODS_TARGET_S = "U"


def eprint(*args: object) -> None:
    print(*args, file=sys.stderr)


def col_to_index(col: str) -> int:
    col = col.strip().upper()
    result = 0
    for ch in col:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_cell_text_for_ods(value: object) -> str:
    text = normalize_text(value)
    if not text:
        return ""

    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()


def extract_cabinet_name(value: str) -> str:
    """
    Приводит имя шкафа к единому виду.

    Поддерживает оба формата:
    - РП-НБ2.УСО.2.3(2).xlsx -> РП-НБ2.УСО.2.3(2)
    - 8. Шкаф РП-НБ2.УСО.2.3(2) - ТБ.xlsx -> РП-НБ2.УСО.2.3(2)
    """
    value = normalize_text(value)
    if not value:
        return ""

    # Если пришло имя файла с расширением — убираем расширение.
    if value.lower().endswith((".xlsx", ".xlsm", ".xls")):
        value = Path(value).stem

    value = re.sub(r"\s+", " ", value).strip()

    # Формат вида: "8. Шкаф РП-НБ2.УСО.2.3(2) - ТБ"
    match = re.search(r"шкаф\s+(.+)$", value, re.IGNORECASE)
    if match:
        value = match.group(1).strip()
    else:
        # Резерв: убираем ведущий номер вида "8. "
        value = re.sub(r"^\d+\s*[.)-]\s*", "", value).strip()

    # Убираем хвост вида " - ТБ" / " – ТБ" / " — ТБ"
    value = re.split(r"\s+[-–—]\s+", value, maxsplit=1)[0].strip()

    # Резерв: если хвост записан без дефиса, например "... ТБ"
    value = re.sub(r"\s+ТБ$", "", value, flags=re.IGNORECASE).strip()

    return value


def normalize_file_key(value: str) -> str:
    cabinet_name = extract_cabinet_name(value)
    return cabinet_name.casefold()


def display_cabinet_name(value: str) -> str:
    cabinet_name = extract_cabinet_name(value)
    return cabinet_name if cabinet_name else normalize_text(value)


def normalize_similar_letters(value: str) -> str:
    value = normalize_text(value).upper()
    table = str.maketrans({
        "А": "A",
        "В": "B",
        "Е": "E",
        "К": "K",
        "М": "M",
        "Н": "H",
        "О": "O",
        "Р": "P",
        "С": "C",
        "Т": "T",
        "У": "Y",
        "Х": "X",
    })
    return value.translate(table)


def normalize_tag_value_for_r(value: str) -> str:
    return normalize_similar_letters(value)


def extract_digits(value: str) -> str:
    value = normalize_text(value)
    m = re.search(r"(\d+)", value)
    return m.group(1) if m else ""


def format_module_name(basket: str, module_slot: str) -> str:
    basket = normalize_text(basket)
    module_slot = normalize_text(module_slot)
    if basket and module_slot:
        return f"{basket}.{module_slot}"
    if basket:
        return basket
    return module_slot


def extract_signal_kind_from_ods(value: str) -> str:
    value = normalize_text(value)
    if not value:
        return ""
    if value.lower() == "iddiag":
        return "DIAG"
    m = re.search(r"id\s*([A-Za-z]+)", value, re.IGNORECASE)
    if m:
        return m.group(1).upper()
    return value.upper()


def extract_signal_kind_from_xlsx(type_value: str, channel_value: str) -> str:
    type_value = normalize_text(type_value).upper()
    channel_value = normalize_text(channel_value).upper()

    if not type_value:
        return ""

    m = re.match(r"([A-Z]+)", type_value)
    base_kind = m.group(1) if m else ""

    if base_kind != "DM":
        return base_kind

    if "OUT" in channel_value:
        return "DO"
    if "IN" in channel_value:
        return "DI"
    if re.fullmatch(r"\d+", channel_value):
        return "DI"
    if re.search(r"\d+", channel_value):
        return "DI"

    return ""


def get_cell_text(cell) -> str:
    parts: List[str] = []
    for p in cell.getElementsByType(P):
        for node in p.childNodes:
            if hasattr(node, "data"):
                parts.append(node.data)
    return "".join(parts).strip()


def get_row_values(row: TableRow, max_col: int) -> List[str]:
    values: List[str] = []

    for node in row.childNodes:
        qname = getattr(node, "qname", None)
        if qname is None:
            continue
        local_name = qname[1]
        if local_name not in ("table-cell", "covered-table-cell"):
            continue

        repeat = int(node.getAttribute("numbercolumnsrepeated") or 1)
        text = get_cell_text(node) if local_name == "table-cell" else ""

        for _ in range(repeat):
            values.append(text)
            if len(values) >= max_col:
                return values

    while len(values) < max_col:
        values.append("")
    return values


def append_empty_cells_until(row: TableRow, target_col: int) -> None:
    current = 0
    for node in row.childNodes:
        qname = getattr(node, "qname", None)
        if qname is None:
            continue
        if qname[1] not in ("table-cell", "covered-table-cell"):
            continue
        current += int(node.getAttribute("numbercolumnsrepeated") or 1)

    while current < target_col:
        row.addElement(TableCell())
        current += 1


def safe_get_attr(node, attr: str):
    try:
        return node.getAttribute(attr)
    except Exception:
        return None


def make_plain_cell_from_node(node):
    new_cell = TableCell()

    for attr in [
        "stylename",
        "valuetype",
        "value",
        "stringvalue",
        "formula",
        "currency",
        "datevalue",
        "timevalue",
        "booleanvalue",
    ]:
        val = safe_get_attr(node, attr)
        if val is not None:
            try:
                new_cell.setAttribute(attr, val)
            except Exception:
                pass

    text = get_cell_text(node)
    if text and safe_get_attr(new_cell, "stringvalue") is None:
        try:
            new_cell.setAttribute("valuetype", "string")
            new_cell.setAttribute("stringvalue", text)
        except Exception:
            pass
        new_cell.addElement(P(text=text))

    return new_cell


def replace_node_with_nodes(row: TableRow, old_node, new_nodes: List):
    ref = old_node.nextSibling
    row.removeChild(old_node)
    for new_node in new_nodes:
        if ref is not None:
            row.insertBefore(new_node, ref)
        else:
            row.addElement(new_node)


def expand_repeated_cell_for_target(row: TableRow, target_col: int):
    logical = 1

    for node in list(row.childNodes):
        qname = getattr(node, "qname", None)
        if qname is None:
            continue
        local_name = qname[1]
        if local_name not in ("table-cell", "covered-table-cell"):
            continue

        repeat = int(node.getAttribute("numbercolumnsrepeated") or 1)
        start = logical
        end = logical + repeat - 1

        if start <= target_col <= end:
            if repeat == 1:
                return

            if local_name == "table-cell":
                new_nodes = [make_plain_cell_from_node(node) for _ in range(repeat)]
            else:
                new_nodes = [TableCell() for _ in range(repeat)]

            replace_node_with_nodes(row, node, new_nodes)
            return

        logical = end + 1


def get_cell_for_write(row: TableRow, target_col: int) -> TableCell:
    append_empty_cells_until(row, target_col)
    expand_repeated_cell_for_target(row, target_col)

    logical = 1
    for node in row.childNodes:
        qname = getattr(node, "qname", None)
        if qname is None:
            continue

        local_name = qname[1]
        if local_name not in ("table-cell", "covered-table-cell"):
            continue

        repeat = int(node.getAttribute("numbercolumnsrepeated") or 1)
        start = logical
        end = logical + repeat - 1

        if start <= target_col <= end:
            if local_name == "covered-table-cell":
                plain = TableCell()
                replace_node_with_nodes(row, node, [plain])
                return plain

            if repeat != 1:
                raise RuntimeError(f"Не удалось развернуть repeated-ячейку для колонки {target_col}.")
            return node

        logical = end + 1

    raise RuntimeError(f"Не удалось получить ячейку для записи в колонку {target_col}")


def clear_children(cell: TableCell) -> None:
    for child in list(cell.childNodes):
        cell.removeChild(child)


def set_cell_text(cell: TableCell, value: str) -> None:
    value = normalize_text(value)

    for attr in [
        "formula", "valuetype", "value", "datevalue", "timevalue",
        "booleanvalue", "stringvalue", "currency",
    ]:
        try:
            if cell.getAttribute(attr) is not None:
                cell.removeAttribute(attr)
        except Exception:
            pass

    clear_children(cell)

    if value:
        cell.setAttribute("valuetype", "string")
        cell.setAttribute("stringvalue", value)
        cell.addElement(P(text=value))


def build_xlsx_index(base_dir: Path) -> Tuple[
    Dict[Tuple[str, str, str, str, str], Tuple[str, str, str]],
    Dict[str, int],
    Set[Tuple[str, str, str]],
    Dict[str, str],
    Dict[Tuple[str, str, str, str, str], Dict[str, object]],
]:
    index: Dict[Tuple[str, str, str, str, str], Tuple[str, str, str]] = {}
    stats = defaultdict(int)
    xlsx_modules: Set[Tuple[str, str, str]] = set()
    cabinet_display_names: Dict[str, str] = {}
    duplicate_details: Dict[Tuple[str, str, str, str, str], Dict[str, object]] = {}
    key_first_rows: Dict[Tuple[str, str, str, str, str], int] = {}

    xlsx_files = sorted(p for p in base_dir.rglob("*.xlsx") if not p.name.startswith("~$"))
    if not xlsx_files:
        raise FileNotFoundError("В папке не найдено ни одного .xlsx файла")

    for xlsx_path in xlsx_files:
        stats["xlsx_files_total"] += 1
        file_key = normalize_file_key(xlsx_path.name)
        cabinet_display_names[file_key] = display_cabinet_name(xlsx_path.name)

        try:
            wb = load_workbook(xlsx_path, read_only=True, data_only=True)
        except Exception as exc:
            eprint(f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось открыть {xlsx_path.name}: {exc}")
            stats["xlsx_open_errors"] += 1
            continue

        if SHEET_NAME_XLSX not in wb.sheetnames:
            stats["xlsx_without_print_sheet"] += 1
            wb.close()
            continue

        ws = wb[SHEET_NAME_XLSX]
        stats["xlsx_with_print_sheet"] += 1

        for row_number, row in enumerate(ws.iter_rows(values_only=True), start=1):
            val_c = normalize_text(row[2] if len(row) > 2 else "")
            val_d = normalize_text(row[3] if len(row) > 3 else "")
            val_e = normalize_text(row[4] if len(row) > 4 else "")
            val_f = normalize_text(row[5] if len(row) > 5 else "")
            val_g = normalize_text(row[6] if len(row) > 6 else "")
            val_h = normalize_text(row[7] if len(row) > 7 else "")
            val_m = normalize_text(row[12] if len(row) > 12 else "")

            basket = normalize_similar_letters(val_f)
            module_slot = extract_digits(val_g)
            channel_num = extract_digits(val_h)
            signal_kind = extract_signal_kind_from_xlsx(val_e, val_h)

            if basket and module_slot:
                xlsx_modules.add((file_key, basket, module_slot))

            if not basket or not module_slot or not channel_num or not signal_kind:
                continue

            key = (file_key, basket, module_slot, channel_num, signal_kind)

            if key in index:
                stats["duplicate_keys"] += 1

                if key not in duplicate_details:
                    first_row = key_first_rows.get(key)
                    rows = []
                    if first_row is not None:
                        rows.append(first_row)
                    rows.append(row_number)

                    duplicate_details[key] = {
                        "file_key": file_key,
                        "cabinet_name": xlsx_path.stem,
                        "basket": basket,
                        "module_slot": module_slot,
                        "channel": channel_num,
                        "signal_kind": signal_kind,
                        "rows": rows,
                    }
                else:
                    duplicate_details[key]["rows"].append(row_number)

                continue

            index[key] = (val_m, val_c, val_d)
            key_first_rows[key] = row_number
            stats["index_rows"] += 1

        wb.close()

    return index, stats, xlsx_modules, cabinet_display_names, duplicate_details


def collect_ods_modules(
    rows: List[TableRow],
    idx_c: int,
    idx_e: int,
    idx_j: int,
    max_read_col: int,
) -> Set[Tuple[str, str, str]]:
    ods_modules: Set[Tuple[str, str, str]] = set()

    current_file_key = ""
    current_basket = ""
    current_module_slot = ""

    logical_row = 0

    for row in rows:
        row_repeat = int(row.getAttribute("numberrowsrepeated") or 1)
        values = get_row_values(row, max_read_col)

        if row_repeat != 1:
            if not any(values):
                break
            raise RuntimeError(
                f"Начиная с {START_ROW}-й строки встретилась непустая повторяющаяся строка "
                f"(логическая строка {logical_row + 1}, repeat={row_repeat})."
            )

        logical_row += 1
        if logical_row < START_ROW:
            continue

        raw_c = values[idx_c - 1] if idx_c <= len(values) else ""
        raw_e = values[idx_e - 1] if idx_e <= len(values) else ""
        raw_j = values[idx_j - 1] if idx_j <= len(values) else ""

        if raw_j:
            current_file_key = normalize_file_key(raw_j)
        if raw_e:
            current_basket = normalize_similar_letters(raw_e)
        if raw_c:
            current_module_slot = extract_digits(raw_c)

        if current_file_key and current_basket and current_module_slot:
            ods_modules.add((current_file_key, current_basket, current_module_slot))

    return ods_modules


def fill_ods(ods_path: Path, create_backup: bool = True) -> Dict[str, object]:
    if not ods_path.exists():
        raise FileNotFoundError(f"ODS файл не найден: {ods_path}")

    base_dir = ods_path.parent
    xlsx_index, build_stats, xlsx_modules, cabinet_display_names, duplicate_details = build_xlsx_index(base_dir)

    doc = load(str(ods_path))
    tables = [t for t in doc.spreadsheet.getElementsByType(Table) if t.getAttribute("name") == SHEET_NAME_ODS]
    if not tables:
        raise RuntimeError(f"В ODS не найден лист {SHEET_NAME_ODS!r}")
    sheet = tables[0]

    rows = sheet.getElementsByType(TableRow)

    idx_c = col_to_index(COL_ODS_MODULE_SLOT)
    idx_e = col_to_index(COL_ODS_BASKET)
    idx_j = col_to_index(COL_ODS_SOURCE_FILE)
    idx_q = col_to_index(COL_ODS_CHANNEL)
    idx_t = col_to_index(COL_ODS_SIGNAL_TYPE)

    idx_p = col_to_index(COL_ODS_TARGET_P)
    idx_r = col_to_index(COL_ODS_TARGET_R)
    idx_s = col_to_index(COL_ODS_TARGET_S)

    max_read_col = max(idx_c, idx_e, idx_j, idx_q, idx_t, idx_p, idx_r, idx_s)

    ods_modules = collect_ods_modules(rows, idx_c, idx_e, idx_j, max_read_col)
    missing_modules = sorted(
        xlsx_modules - ods_modules,
        key=lambda x: (x[0], x[1], int(x[2]) if x[2].isdigit() else x[2]),
    )

    stats: Dict[str, object] = defaultdict(int)
    stats.update(build_stats)
    stats["xlsx_modules_total"] = len(xlsx_modules)
    stats["ods_modules_total"] = len(ods_modules)
    stats["missing_modules"] = missing_modules
    stats["cabinet_display_names"] = cabinet_display_names
    stats["duplicate_details"] = duplicate_details

    current_file_key = ""
    current_basket = ""
    current_module_slot = ""

    logical_row = 0

    for row in rows:
        row_repeat = int(row.getAttribute("numberrowsrepeated") or 1)
        values = get_row_values(row, max_read_col)

        if row_repeat != 1:
            if not any(values):
                break
            raise RuntimeError(
                f"Начиная с {START_ROW}-й строки встретилась непустая повторяющаяся строка "
                f"(логическая строка {logical_row + 1}, repeat={row_repeat})."
            )

        logical_row += 1
        if logical_row < START_ROW:
            continue

        raw_c = values[idx_c - 1] if idx_c <= len(values) else ""
        raw_e = values[idx_e - 1] if idx_e <= len(values) else ""
        raw_j = values[idx_j - 1] if idx_j <= len(values) else ""
        raw_q = values[idx_q - 1] if idx_q <= len(values) else ""
        raw_t = values[idx_t - 1] if idx_t <= len(values) else ""

        if raw_j:
            current_file_key = normalize_file_key(raw_j)
        if raw_e:
            current_basket = normalize_similar_letters(raw_e)
        if raw_c:
            current_module_slot = extract_digits(raw_c)

        signal_kind = extract_signal_kind_from_ods(raw_t)
        if signal_kind == "DIAG":
            stats["rows_skipped_diag"] += 1
            continue

        channel_num = extract_digits(raw_q)

        if not signal_kind:
            stats["rows_skipped_no_type"] += 1
            continue
        if not current_file_key:
            stats["rows_skipped_no_file"] += 1
            continue
        if not current_basket or not current_module_slot:
            stats["rows_skipped_no_module"] += 1
            continue
        if not channel_num:
            stats["rows_skipped_no_channel"] += 1
            continue

        key = (current_file_key, current_basket, current_module_slot, channel_num, signal_kind)
        found = xlsx_index.get(key)

        if not found:
            stats["rows_not_found"] += 1
            continue

        value_p, value_r, value_s = found

        value_p = normalize_cell_text_for_ods(value_p)
        value_r = normalize_tag_value_for_r(normalize_cell_text_for_ods(value_r))
        value_s = normalize_cell_text_for_ods(value_s)

        set_cell_text(get_cell_for_write(row, idx_p), value_p)
        set_cell_text(get_cell_for_write(row, idx_r), value_r)
        set_cell_text(get_cell_for_write(row, idx_s), value_s)
        stats["rows_filled"] += 1

    if create_backup:
        backup_path = ods_path.with_name(f"{ods_path.stem}.backup_before_fill{ods_path.suffix}")
        shutil.copy2(ods_path, backup_path)
        stats["backup_created"] = 1

    doc.save(str(ods_path))
    stats["saved"] = 1
    return stats


def print_missing_modules(stats: Dict[str, object], structured_output: bool = False) -> None:
    missing_modules = stats.get("missing_modules", [])
    cabinet_display_names = stats.get("cabinet_display_names", {})

    if structured_output:
        print(f"NEW_MODULES_COUNT: {len(missing_modules)}")
        if missing_modules:
            print("Список новых модулей:")
            for file_key, basket, module_slot in missing_modules:
                cabinet_name = cabinet_display_names.get(file_key, file_key)
                module_name = format_module_name(basket, module_slot)
                print(f"NEW_MODULE: {cabinet_name} | {module_name}")
        return

    if not missing_modules:
        print("Новых модулей в XLSX, отсутствующих в ODS: нет")
        return

    print(f"Новых модулей в XLSX, отсутствующих в ODS: {len(missing_modules)}")
    print("Список новых модулей:")
    for file_key, basket, module_slot in missing_modules:
        cabinet_name = cabinet_display_names.get(file_key, file_key)
        module_name = format_module_name(basket, module_slot)
        print(f"  - {cabinet_name}: {module_name}")


def print_duplicate_details(stats: Dict[str, object], structured_output: bool = False) -> None:
    duplicate_details = stats.get("duplicate_details", {})

    if structured_output:
        print(f"DUPLICATE_KEYS_COUNT: {len(duplicate_details)}")
        if duplicate_details:
            print("Список дубликатов:")
            for detail in duplicate_details.values():
                cabinet_name = detail.get("cabinet_name", detail.get("file_key", ""))
                basket = str(detail.get("basket", ""))
                module_slot = str(detail.get("module_slot", ""))
                module_name = format_module_name(basket, module_slot)
                channel = detail.get("channel", "")
                signal_kind = detail.get("signal_kind", "")
                rows = detail.get("rows", [])
                rows_text = ", ".join(str(x) for x in rows)
                print(
                    f"DUPLICATE_KEY: {cabinet_name} | {module_name} | "
                    f"{channel} | {signal_kind} | {rows_text}"
                )
        return

    if not duplicate_details:
        print("Дубликатов ключей в XLSX: нет")
        return

    print(f"Дубликатов ключей в XLSX: {len(duplicate_details)}")
    print("Список дубликатов:")
    for detail in duplicate_details.values():
        cabinet_name = detail.get("cabinet_name", detail.get("file_key", ""))
        basket = str(detail.get("basket", ""))
        module_slot = str(detail.get("module_slot", ""))
        module_name = format_module_name(basket, module_slot)
        channel = detail.get("channel", "")
        signal_kind = detail.get("signal_kind", "")
        rows = detail.get("rows", [])
        rows_text = ", ".join(str(x) for x in rows)
        print(
            f"  - {cabinet_name}: {module_name}, канал {channel}, "
            f"тип {signal_kind}, строки XLSX: {rows_text}"
        )


def main() -> int:
    parser = argparse.ArgumentParser(description="Заполнение листа 'Сигналы' в ODS по XLSX-файлам с листом 'Печать'.")
    parser.add_argument("ods_path", help="Путь к ODS-файлу")
    parser.add_argument("--no-backup", action="store_true", help="Не создавать резервную копию")
    parser.add_argument(
        "--structured-output",
        action="store_true",
        help="Выводить служебные маркеры для PowerShell-оболочки",
    )
    args = parser.parse_args()

    ods_path = Path(args.ods_path).resolve()

    try:
        stats = fill_ods(ods_path, create_backup=not args.no_backup)
    except Exception as exc:
        eprint(f"[ОШИБКА] {exc}")
        return 1

    print("Готово.")
    print(f"ODS: {ods_path}")
    print_missing_modules(stats, structured_output=args.structured_output)
    print_duplicate_details(stats, structured_output=args.structured_output)
    print(f"Модулей в XLSX: {stats.get('xlsx_modules_total', 0)}")
    print(f"Модулей в ODS: {stats.get('ods_modules_total', 0)}")
    print(f"Заполнено строк: {stats.get('rows_filled', 0)}")
    print(f"Не найдено совпадений: {stats.get('rows_not_found', 0)}")
    print(f"Пропущено idDiag: {stats.get('rows_skipped_diag', 0)}")
    print(f"Пропущено без типа: {stats.get('rows_skipped_no_type', 0)}")
    print(f"Пропущено без файла: {stats.get('rows_skipped_no_file', 0)}")
    print(f"Пропущено без модуля: {stats.get('rows_skipped_no_module', 0)}")
    print(f"Пропущено без канала: {stats.get('rows_skipped_no_channel', 0)}")
    print(f"XLSX найдено: {stats.get('xlsx_files_total', 0)}")
    print(f"XLSX с листом 'Печать': {stats.get('xlsx_with_print_sheet', 0)}")
    print(f"Строк в индексе: {stats.get('index_rows', 0)}")
    if stats.get('duplicate_keys', 0):
        print(f"Дубликатов ключей в XLSX пропущено: {stats.get('duplicate_keys', 0)}")
    if stats.get('backup_created', 0):
        print("Резервная копия создана рядом с ODS.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
