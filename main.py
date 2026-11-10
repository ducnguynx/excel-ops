import os
import re
import shutil
import copy
import openpyxl
from openpyxl.cell import Cell
from openpyxl.reader.excel import ExcelReader
from openpyxl.worksheet._reader import WorksheetReader
from openpyxl.workbook.defined_name import DefinedNameList
import xlrd
from thefuzz import fuzz


openpyxl.workbook.defined_name.DefinedNameList.from_tree = classmethod(
    lambda cls, node: DefinedNameList()
)
xlrd.Book.names_epilogue = lambda self: None


def load_summary_xlsx_sheet(filepath):
    reader = ExcelReader(filepath, read_only=False, keep_vba=False, data_only=True)
    try:
        reader.read_manifest()
        reader.read_workbook()
        for sheet, rel in reader.parser.find_sheets():
            if (
                rel.target in reader.valid_files
                and "chartsheet" not in rel.Type
                and sheet.state in (None, "visible")
                and "summary" in sheet.name.lower()
            ):
                reader.read_strings()
                ws = reader.wb.create_sheet(sheet.name)
                with reader.archive.open(rel.target) as fh:
                    ws_parser = WorksheetReader(
                        ws, fh, reader.shared_strings, reader.data_only, reader.rich_text
                    )
                    for _, row in ws_parser.parser.parse():
                        for cell in row:
                            c = Cell(ws, row=cell["row"], column=cell["column"])
                            c._value = cell["value"]
                            c.data_type = cell["data_type"]
                            ws._cells[(cell["row"], cell["column"])] = c
                    if ws._cells:
                        ws._current_row = ws.max_row
                    ws_parser.bind_merged_cells()
                return ws
        return None
    finally:
        reader.archive.close()


def load_summary_xls_sheet(filepath):
    wb = xlrd.open_workbook(filepath, formatting_info=True, on_demand=True)
    try:
        for sheet_name in wb.sheet_names():
            if "summary" in sheet_name.lower():
                return wb.sheet_by_name(sheet_name)
        return None
    finally:
        wb.release_resources()

SUPPORTED_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}

DESTINATION_COLUMNS = {
    "Advance payment": 7,
    "Total before VAT": 8,
    "Remeasured Work": 9,
    "VAT": 10,
    "20%": 11,
    "Retention": 12,
    "Advance payment recovery": 13,
    "Remaining": 14,
    "Total payable": 16,
}

FORMULA_REMAINING_COL = 15
CROSSCHECK_COL = 17

FIELD_SEARCH_TERMS = {
    "Advance payment": ("Advance payment",),
    "Total before VAT": ("Total before VAT",),
    "Remeasured Work": ("Remeasured Work",),
    "VAT": ("VAT (10%)", "VAT"),
    "20%": ("Pay 20% of the previous value",),
    "Retention": ("Retention",),
    "Advance payment recovery": (
        "Advance Payment Recovery",
        "Refund of Advance Payment",
    ),
    "Remaining": ("Remaining",),
    "Total payable": (
        "Total payable",
        "Total amount",
        "Net Payable Amt",
    ),
}

ABSOLUTE_VALUE_FIELDS = {"Advance payment recovery", "VAT"}

RETENTION_FIELDS = {"Retention"}

CONTRACTOR_NAME_ALIASES = {
    "FT": "FOUNTECH",
    "FOUTECH": "FOUNTECH",
    "LIFFT": "LIFTECH",
    "LIFFTECH": "LIFTECH",
    "LIFFTEC": "LIFTECH",
    "BT": "BAO TRAN",
    "GLONG": "GIANG LONG",
    "ORIENTAL": "ORIEN",
    "SH": "SONG HONG",
    "TOWRELAND": "TOWERLAND",
    "VT": "VITECONS",
    "VITEC": "VITECONS",
}


def normalize_text(value):
    return " ".join(str(value).upper().split())


def is_xls_file(filepath):
    return os.path.splitext(filepath)[1].lower() == ".xls"


def iter_sheet_cells(sheet, xls_format, max_row=None, max_col=20):
    if xls_format:
        row_count = sheet.nrows if max_row is None else min(sheet.nrows, max_row)
        col_count = sheet.ncols if max_col is None else min(sheet.ncols, max_col)
        for row_idx in range(row_count):
            for col_idx in range(col_count):
                yield row_idx + 1, col_idx + 1, sheet.cell_value(row_idx, col_idx)
        return

    for row in sheet.iter_rows(max_row=max_row, max_col=max_col):
        for cell in row:
            yield cell.row, cell.column, cell.value


def get_cell_value(sheet, row, column, xls_format):
    if xls_format:
        return sheet.cell_value(row - 1, column - 1)
    return sheet.cell(row=row, column=column).value


def get_merged_col_span(sheet, row, column, xls_format):
    if xls_format:
        for rlo, rhi, clo, chi in sheet.merged_cells:
            if rlo <= row - 1 < rhi and clo <= column - 1 < chi:
                return chi - clo
        return 1

    for merged_range in sheet.merged_cells.ranges:
        if (
            merged_range.min_row <= row <= merged_range.max_row
            and merged_range.min_col <= column <= merged_range.max_col
        ):
            return merged_range.max_col - merged_range.min_col + 1
    return 1


def ensure_master_columns(master_sheet):
    col_j_header = normalize_text(master_sheet.cell(row=2, column=10).value or "")
    if col_j_header != "VAT":
        master_sheet.insert_cols(10, 2)
        for col_idx in (10, 11):
            src_cell = master_sheet.cell(row=2, column=9)
            dst_cell = master_sheet.cell(row=2, column=col_idx)
            if src_cell.has_style:
                dst_cell.font = copy.copy(src_cell.font)
                dst_cell.fill = copy.copy(src_cell.fill)
                dst_cell.border = copy.copy(src_cell.border)
                dst_cell.alignment = copy.copy(src_cell.alignment)
        for row_idx in range(3, master_sheet.max_row + 1):
            formula_cell = master_sheet.cell(row=row_idx, column=FORMULA_REMAINING_COL)
            if isinstance(formula_cell.value, str) and formula_cell.value.startswith("="):
                formula_cell.value = (
                    f"=I{row_idx}-J{row_idx}+K{row_idx}"
                    f"-L{row_idx}-M{row_idx}-N{row_idx}"
                )

    master_sheet.cell(row=2, column=DESTINATION_COLUMNS["VAT"], value="VAT")
    master_sheet.cell(row=2, column=DESTINATION_COLUMNS["20%"], value="20%")
    master_sheet.cell(
        row=2, column=DESTINATION_COLUMNS["Total payable"], value="Total payable"
    )
    master_sheet.cell(row=2, column=CROSSCHECK_COL, value="Kiểm tra")


def find_round_start_row(master_sheet, round_name):
    last_data_row = 2
    round_rows = []
    for row_idx in range(3, master_sheet.max_row + 1):
        round_val = master_sheet.cell(row=row_idx, column=3).value
        contractor_val = master_sheet.cell(row=row_idx, column=4).value
        if has_value(round_val) or has_value(contractor_val):
            last_data_row = row_idx
            if normalize_text(round_val) == normalize_text(round_name):
                round_rows.append(row_idx)

    if round_rows:
        for row_idx in round_rows:
            for col_idx in range(2, CROSSCHECK_COL + 1):
                if col_idx in (5, FORMULA_REMAINING_COL):
                    continue
                master_sheet.cell(row=row_idx, column=col_idx, value=None)
        return round_rows[0]

    return last_data_row + 1


def get_round_name(source_folder):
    folder_name = os.path.basename(os.path.normpath(source_folder))
    round_number = re.search(r"\d+", folder_name)
    return f"Đợt {round_number.group()}" if round_number else folder_name


def get_contractor_name(filepath, source_folder):
    relative_filepath = os.path.relpath(filepath, source_folder)
    folder_name = normalize_text(relative_filepath.split(os.sep)[0])
    folder_name = re.sub(r"[-_\s]*\d+$", "", folder_name).strip()
    return CONTRACTOR_NAME_ALIASES.get(folder_name, folder_name)


def fuzzy_prefix_match(search_desc, cell_text):
    normalized_search = normalize_text(search_desc)
    candidate_words = normalize_text(cell_text).split()
    search_word_count = len(normalized_search.split())
    candidate_prefix = " ".join(candidate_words[:search_word_count])
    return fuzz.ratio(normalized_search, candidate_prefix) >= 80


def matches_field(field_name, cell_text, item_code):
    normalized_text = normalize_text(cell_text)
    normalized_code = normalize_text(item_code)

    if field_name == "Advance payment":
        if "RECOVERY" in normalized_text or "REFUND" in normalized_text:
            return False
        matched = any(
            fuzzy_prefix_match(search_term, cell_text)
            for search_term in FIELD_SEARCH_TERMS[field_name]
        )
        return matched and normalized_code.startswith("B")

    if field_name == "20%":
        return (
            normalized_code.startswith("E")
            and "20%" in normalized_text
            and "PREVIOUS" in normalized_text
        )

    matched = any(
        fuzzy_prefix_match(search_term, cell_text)
        for search_term in FIELD_SEARCH_TERMS[field_name]
    )

    if field_name == "Total before VAT":
        return matched and normalized_code.startswith("C")

    if field_name in RETENTION_FIELDS | {"Advance payment recovery", "VAT"}:
        return matched and normalized_code.startswith("D")

    return matched


def get_output_value(field_name, value):
    if field_name in ABSOLUTE_VALUE_FIELDS and isinstance(value, (int, float)):
        return abs(value)
    return value


def has_value(value):
    return value not in (None, "")


def get_ipc_number(sheet, xls_format):
    for row, column, value in iter_sheet_cells(sheet, xls_format, max_row=10):
        if not isinstance(value, str):
            continue

        normalized_value = normalize_text(value)
        if "CERTIFICATION" not in normalized_value or "NO" not in normalized_value:
            continue

        certificate_number = get_cell_value(sheet, row, column + 1, xls_format)
        if certificate_number in (None, ""):
            continue
        if isinstance(certificate_number, float) and certificate_number.is_integer():
            certificate_number = int(certificate_number)
        return f"IPC{certificate_number}"

    return None


def extract_intersection_data(source_folder, master_filepath, master_wb=None):
    # 1. DEFINE YOUR MAPPING HERE
    # Keys = The description text to search for in the rows (e.g., Column B text)
    # Values = The exact Column Header name in your Master Sheet where the data should go
    field_names = tuple(DESTINATION_COLUMNS)
    round_name = get_round_name(source_folder)
    serial_number = 0

    # Load Master File (Specifically "Sheet1")
    owns_workbook = master_wb is None
    if owns_workbook:
        master_wb = openpyxl.load_workbook(master_filepath)
    if "Sheet1" not in master_wb.sheetnames:
        raise ValueError("Master file does not contain 'Sheet1'")
    master_sheet = master_wb["Sheet1"]

    ensure_master_columns(master_sheet)
    current_row = find_round_start_row(master_sheet, round_name)

    candidate_files = []
    for root, _, filenames in os.walk(source_folder):
        for filename in sorted(filenames):
            extension = os.path.splitext(filename)[1].lower()
            if extension not in SUPPORTED_EXTENSIONS or filename.startswith(("~", ".")):
                continue
            filepath = os.path.join(root, filename)
            contractor_name = get_contractor_name(filepath, source_folder)
            candidate_files.append((contractor_name, filepath))

    candidate_files.sort(key=lambda item: (item[0], item[1]))

    for contractor_name, filepath in candidate_files:
        print(f"Processing: {filepath}...")
        xls_format = is_xls_file(filepath)

        # 2. FIND THE FIRST SHEET CONTAINING "SUMMARY"
        target_sheet = (
            load_summary_xls_sheet(filepath)
            if xls_format
            else load_summary_xlsx_sheet(filepath)
        )

        if not target_sheet:
            print("  -> Skipped: No 'Summary' sheet found.")
            continue

        # 3. FIND THE "THIS MONTH" COLUMN (X-Axis)
        target_row_idx = None
        target_col_idx = None
        for row, column, value in iter_sheet_cells(target_sheet, xls_format, max_row=50):
            if isinstance(value, str) and "THIS MONTH" in normalize_text(value):
                target_row_idx = row
                target_col_idx = column
                break

        if not target_col_idx:
            print("  -> Skipped: Could not find 'THIS MONTH' column.")
            continue

        # 4. FIND THE ROWS FOR THE FIELDS (Y-Axis) & EXTRACT THE INTERSECTION
        matched_values = {field_name: [] for field_name in field_names}
        anomaly_fields = set()
        ipc_number = get_ipc_number(target_sheet, xls_format)

        sheet_level_anomaly = (
            get_merged_col_span(target_sheet, target_row_idx, target_col_idx, xls_format) > 1
        )

        for row, column, value in iter_sheet_cells(target_sheet, xls_format, max_row=100):
            if not isinstance(value, str):
                continue

            item_code = (
                get_cell_value(target_sheet, row, column - 1, xls_format)
                if column > 1
                else None
            )
            if not has_value(item_code):
                item_code = get_cell_value(target_sheet, row, 1, xls_format)

            for field_name in field_names:
                if matches_field(field_name, value, item_code):
                    value_col_idx = (
                        target_col_idx - 1
                        if field_name == "Advance payment"
                        else target_col_idx
                    )
                    cell_value = get_cell_value(
                        target_sheet,
                        row,
                        value_col_idx,
                        xls_format,
                    )
                    if has_value(cell_value):
                        matched_values[field_name].append(cell_value)

        extracted_data = {}
        for field_name, values in matched_values.items():
            if field_name in RETENTION_FIELDS:
                numeric_values = [
                    v for v in values if isinstance(v, (int, float))
                ]
                if len(numeric_values) != len(values):
                    anomaly_fields.add(field_name)
                elif numeric_values:
                    extracted_data[field_name] = abs(sum(numeric_values))
                continue

            if len(values) > 1:
                anomaly_fields.add(field_name)
            if values:
                extracted_data[field_name] = get_output_value(
                    field_name, values[-1]
                )

        if not extracted_data or sheet_level_anomaly:
            anomaly_fields.update(field_names)

        # 5. WRITE TO TARGET SHEET
        serial_number += 1
        next_row = current_row
        current_row += 1
        master_sheet.cell(row=next_row, column=2, value=serial_number)
        master_sheet.cell(row=next_row, column=3, value=round_name)
        master_sheet.cell(
            row=next_row,
            column=4,
            value=contractor_name,
        )
        master_sheet.cell(row=next_row, column=6, value=ipc_number)

        master_sheet.cell(
            row=next_row,
            column=FORMULA_REMAINING_COL,
            value=(
                f"=I{next_row}-J{next_row}+K{next_row}"
                f"-L{next_row}-M{next_row}-N{next_row}"
            ),
        )

        master_sheet.cell(
            row=next_row,
            column=CROSSCHECK_COL,
            value=f'=IF(AND(ISNUMBER(O{next_row}),ISNUMBER(P{next_row}),ROUND(O{next_row},0)=ROUND(P{next_row},0)),"OK","check")',
        )

        if sheet_level_anomaly or not extracted_data:
            for col_idx in set(DESTINATION_COLUMNS.values()):
                master_sheet.cell(row=next_row, column=col_idx, value="check")
            continue

        for field_name, value in extracted_data.items():
            if field_name in anomaly_fields:
                value = "check"
            master_sheet.cell(
                row=next_row,
                column=DESTINATION_COLUMNS[field_name],
                value=value,
            )

    if owns_workbook:
        master_wb.save(master_filepath)
        print(f"\nSuccess! Data compiled into {master_filepath}")


if __name__ == "__main__":
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    SOURCE_DIRS = sorted(
        [
            os.path.join(BASE_DIR, name)
            for name in os.listdir(BASE_DIR)
            if os.path.isdir(os.path.join(BASE_DIR, name))
            and not name.startswith((".", "_"))
        ],
        key=lambda path: (
            int(m.group())
            if (m := re.search(r"\d+", os.path.basename(path)))
            else float("inf"),
            os.path.basename(path).lower(),
        ),
    )
    MASTER_FILE = os.path.join(BASE_DIR, "Danh mục SA 08 - NOVATION.xlsx")
    TEST_MASTER_FILE = os.path.join(BASE_DIR, "Danh mục SA 08 - NOVATION - test.xlsx")

    shutil.copy2(MASTER_FILE, TEST_MASTER_FILE)
    master_workbook = openpyxl.load_workbook(TEST_MASTER_FILE)
    for source_dir in SOURCE_DIRS:
        extract_intersection_data(source_dir, TEST_MASTER_FILE, master_workbook)
    master_workbook.save(TEST_MASTER_FILE)
    print(f"\nSuccess! Data compiled into {TEST_MASTER_FILE}")
