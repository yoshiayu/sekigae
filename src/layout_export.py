from __future__ import annotations

import io
import json
import re
from collections import defaultdict
from datetime import datetime
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.cell import range_boundaries

TEMPLATE_ROWS = 19
TEMPLATE_COLS = 20
_TEMPLATE_TITLE_CELL = (2, 2)
_TITLE_MERGE_RANGE = "B2:T2"
_SCREEN_TOP_CELL = (5, 6)
_SCREEN_BOTTOM_CELL = (19, 6)
_SCREEN_TOP_MERGE_RANGE = "F5:Q5"
_SCREEN_BOTTOM_MERGE_RANGE = "F19:Q19"

# 座席表（テンプレ）の座標に合わせた 13テーブル x 6席 の配置定義
# スクリーン視点で千鳥配置: 前4 / 中5 / 後4
# テーブル番号は以下の順（左→右）
# 1:前1, 2:中1, 3:後1, 4:前2, 5:中2, 6:後2, ... , 13:中5
_FRONT_ROWS = (8, 9, 10)
_MIDDLE_ROWS = (12, 13, 14)
_BACK_ROWS = (16, 17, 18)
_FRONT_COL_PAIRS = ((4, 6), (8, 10), (12, 14), (16, 18))
_MIDDLE_COL_PAIRS = ((2, 4), (6, 8), (10, 12), (14, 16), (18, 20))
_BACK_COL_PAIRS = ((4, 6), (8, 10), (12, 14), (16, 18))
_TABLE_POSITION_ORDER: tuple[tuple[str, int], ...] = (
    ("front", 0),
    ("middle", 0),
    ("back", 0),
    ("front", 1),
    ("middle", 1),
    ("back", 1),
    ("front", 2),
    ("middle", 2),
    ("back", 2),
    ("front", 3),
    ("middle", 3),
    ("back", 3),
    ("middle", 4),
)

_GOOGLE_SCOPES = ("https://www.googleapis.com/auth/spreadsheets",)
_EXCEL_INVALID_TITLE = re.compile(r"[\\/*?:\[\]]")
_GOOGLE_INVALID_TITLE = re.compile(r"[\\/?*\[\]:]")
_FILE_INVALID = re.compile(r"[\\/:*?\"<>|]")
_DEFAULT_TEMPLATE_SHEET_NAME = "座席表（テンプレ）"

_TITLE_FONT = Font(name="Meiryo UI", size=14, bold=True)
_SCREEN_FONT = Font(name="Meiryo UI", size=13, bold=True)
_SEAT_FONT = Font(name="Meiryo UI", size=10, bold=True)
_CENTER = Alignment(horizontal="center", vertical="center")
_CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
_THIN_SIDE = Side(style="thin", color="000000")
_THIN_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)
_TITLE_FILL = PatternFill(fill_type="solid", fgColor="FFF200")
_SCREEN_FILL = PatternFill(fill_type="solid", fgColor="F7F7F7")
_SEAT_FILL = PatternFill(fill_type="solid", fgColor="E9E9E9")


def default_daily_sheet_name(reference_dt: datetime | None = None) -> str:
    dt = reference_dt or datetime.now()
    return f"{dt.month}月{dt.day}日の座席表"


def normalize_base_sheet_name(value: str) -> str:
    text = value.strip()
    return text if text else default_daily_sheet_name()


def default_table_layout_rows(table_count: int = 13) -> list[list[int]]:
    front = [no for no in (1, 4, 7, 10) if no <= table_count]
    middle = [no for no in (2, 5, 8, 11, 13) if no <= table_count]
    back = [no for no in (3, 6, 9, 12) if no <= table_count]
    return [front, middle, back]


def build_layout_payload(
    rows: list[dict[str, Any]],
    base_sheet_name: str,
    include_skill: bool = True,
    table_layout_rows: list[list[int]] | None = None,
) -> dict[str, Any]:
    normal_name = normalize_base_sheet_name(base_sheet_name)
    mirrored_name = f"{normal_name}（反転）"
    return {
        "normal_name": normal_name,
        "mirrored_name": mirrored_name,
        "normal_grid": build_template_layout_grid(
            rows,
            normal_name,
            mirrored=False,
            include_skill=include_skill,
            table_layout_rows=table_layout_rows,
        ),
        "mirrored_grid": build_template_layout_grid(
            rows,
            mirrored_name,
            mirrored=True,
            include_skill=include_skill,
            table_layout_rows=table_layout_rows,
        ),
    }


def safe_download_filename(value: str, suffix: str = "") -> str:
    cleaned = _FILE_INVALID.sub("_", value.strip())
    cleaned = cleaned.strip(" .")
    base = cleaned or "seating_layout"
    return f"{base}{suffix}" if suffix else base


def build_template_layout_grid(
    rows: list[dict[str, Any]],
    sheet_title: str,
    mirrored: bool,
    include_skill: bool = True,
    table_layout_rows: list[list[int]] | None = None,
) -> list[list[str]]:
    grid = [["" for _ in range(TEMPLATE_COLS)] for _ in range(TEMPLATE_ROWS)]
    grid[_TEMPLATE_TITLE_CELL[0] - 1][_TEMPLATE_TITLE_CELL[1] - 1] = sheet_title
    screen_row, screen_col = _SCREEN_BOTTOM_CELL if mirrored else _SCREEN_TOP_CELL
    grid[screen_row - 1][screen_col - 1] = "スクリーン"

    table_slots = _resolve_table_slots(table_layout_rows)
    members_by_table: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        try:
            table_no = int(row["table_no"])
        except (TypeError, ValueError, KeyError):
            continue
        if table_no in table_slots:
            members_by_table[table_no].append(row)

    for table_no in sorted(table_slots.keys()):
        slots = table_slots[table_no]
        members = sorted(members_by_table.get(table_no, []), key=_member_sort_key)
        # mirrored=True は「前から後ろへ見る視点」（左右反転ではなく前後方向を反転）
        active_slots = [_front_to_back_slot(slot) for slot in slots] if mirrored else slots

        for idx, (r_idx, c_idx) in enumerate(active_slots):
            text = (
                _format_member_cell(members[idx], include_skill=include_skill)
                if idx < len(members)
                else ""
            )
            grid[r_idx - 1][c_idx - 1] = text

    return grid


def export_layouts_to_excel_bytes(
    normal_grid: list[list[str]],
    mirrored_grid: list[list[str]],
    normal_sheet_name: str,
    mirrored_sheet_name: str,
) -> bytes:
    workbook = Workbook()
    used_titles: set[str] = set()

    normal_title = _unique_title(
        _sanitize_title(normal_sheet_name, max_len=31, invalid_re=_EXCEL_INVALID_TITLE),
        used_titles,
        max_len=31,
    )
    ws_normal = workbook.active
    ws_normal.title = normal_title
    _write_grid_to_worksheet(ws_normal, normal_grid)
    _style_excel_worksheet(ws_normal, mirrored=False)

    mirrored_title = _unique_title(
        _sanitize_title(mirrored_sheet_name, max_len=31, invalid_re=_EXCEL_INVALID_TITLE),
        used_titles,
        max_len=31,
    )
    ws_mirrored = workbook.create_sheet(title=mirrored_title)
    _write_grid_to_worksheet(ws_mirrored, mirrored_grid)
    _style_excel_worksheet(ws_mirrored, mirrored=True)

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def publish_layouts_to_google_sheets(
    service_account_json_bytes: bytes,
    spreadsheet_ref: str,
    template_sheet_name: str,
    normal_sheet_name: str,
    normal_grid: list[list[str]],
    mirrored_sheet_name: str,
    mirrored_grid: list[list[str]],
) -> dict[str, str]:
    import gspread
    from google.oauth2.service_account import Credentials

    if not service_account_json_bytes:
        raise ValueError("サービスアカウントJSONが空です。")

    try:
        account_info = json.loads(service_account_json_bytes.decode("utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ValueError("サービスアカウントJSONの読み取りに失敗しました。") from exc

    spreadsheet_id = extract_spreadsheet_id(spreadsheet_ref)
    if not spreadsheet_id:
        raise ValueError("出力先スプレッドシートID(URL)を指定してください。")

    sheet_name = template_sheet_name.strip()
    if not sheet_name:
        raise ValueError("テンプレートシート名を指定してください。")

    credentials = Credentials.from_service_account_info(account_info, scopes=list(_GOOGLE_SCOPES))
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(spreadsheet_id)
    available_titles = [ws.title for ws in spreadsheet.worksheets()]
    resolved_sheet_name = _resolve_template_sheet_name(sheet_name, available_titles)
    if not resolved_sheet_name:
        preview = ", ".join(available_titles[:10]) if available_titles else "（シートなし）"
        raise ValueError(
            f"テンプレートシート「{sheet_name}」が見つかりません。"
            f"存在するシート名: {preview}"
        )
    template_ws = spreadsheet.worksheet(resolved_sheet_name)

    existing_titles = {ws.title for ws in spreadsheet.worksheets()}
    normal_title = _unique_title(
        _sanitize_title(normal_sheet_name, max_len=100, invalid_re=_GOOGLE_INVALID_TITLE),
        existing_titles,
        max_len=100,
    )
    mirrored_title = _unique_title(
        _sanitize_title(mirrored_sheet_name, max_len=100, invalid_re=_GOOGLE_INVALID_TITLE),
        existing_titles,
        max_len=100,
    )

    ws_normal = template_ws.duplicate(new_sheet_name=normal_title)
    ws_mirrored = template_ws.duplicate(new_sheet_name=mirrored_title)

    update_range = f"A1:{_to_col_label(TEMPLATE_COLS)}{TEMPLATE_ROWS}"
    ws_normal.update(range_name=update_range, values=normal_grid)
    ws_mirrored.update(range_name=update_range, values=mirrored_grid)

    return {
        "spreadsheet_id": spreadsheet_id,
        "requested_template_sheet_name": sheet_name,
        "used_template_sheet_name": resolved_sheet_name,
        "normal_sheet_name": normal_title,
        "normal_sheet_url": f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit#gid={ws_normal.id}",
        "mirrored_sheet_name": mirrored_title,
        "mirrored_sheet_url": f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit#gid={ws_mirrored.id}",
    }


def extract_spreadsheet_id(value: str) -> str:
    text = value.strip()
    if not text:
        return ""
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", text)
    if match:
        return match.group(1)
    return text


def _build_table_slots() -> dict[int, list[tuple[int, int]]]:
    slots: dict[int, list[tuple[int, int]]] = {}
    zone_layout: dict[str, tuple[tuple[int, ...], tuple[tuple[int, int], ...]]] = {
        "front": (_FRONT_ROWS, _FRONT_COL_PAIRS),
        "middle": (_MIDDLE_ROWS, _MIDDLE_COL_PAIRS),
        "back": (_BACK_ROWS, _BACK_COL_PAIRS),
    }
    for table_no, (zone, idx) in enumerate(_TABLE_POSITION_ORDER, start=1):
        if zone not in zone_layout:
            raise ValueError(f"unknown zone: {zone}")
        row_group, col_pairs = zone_layout[zone]
        left_col, right_col = col_pairs[idx]
        table_slots: list[tuple[int, int]] = []
        for row in row_group:
            table_slots.append((row, left_col))
            table_slots.append((row, right_col))
        slots[table_no] = table_slots
    return slots


def _build_reverse_map(values: list[int]) -> dict[int, int]:
    return {value: values[-(idx + 1)] for idx, value in enumerate(values)}


_TABLE_SLOTS = _build_table_slots()
_SEAT_COORDS = {(row, col) for slots in _TABLE_SLOTS.values() for row, col in slots}
_SEAT_ROWS = sorted({row for row, _ in _SEAT_COORDS})
_SEAT_COLS = sorted({col for _, col in _SEAT_COORDS})
_FRONT_ROW_MAP = _build_reverse_map(_SEAT_ROWS)
_FRONT_COL_MAP = _build_reverse_map(_SEAT_COLS)


def _resolve_table_slots(table_layout_rows: list[list[int]] | None) -> dict[int, list[tuple[int, int]]]:
    if not table_layout_rows:
        return dict(_TABLE_SLOTS)

    custom = _build_table_slots_from_layout_rows(table_layout_rows)
    return custom if custom else dict(_TABLE_SLOTS)


def _build_table_slots_from_layout_rows(
    table_layout_rows: list[list[int]],
) -> dict[int, list[tuple[int, int]]]:
    row_groups: tuple[tuple[int, ...], ...] = (_FRONT_ROWS, _MIDDLE_ROWS, _BACK_ROWS)
    row_col_pairs: tuple[tuple[tuple[int, int], ...], ...] = (
        _FRONT_COL_PAIRS,
        _MIDDLE_COL_PAIRS,
        _BACK_COL_PAIRS,
    )

    slots: dict[int, list[tuple[int, int]]] = {}
    for row_idx in range(min(3, len(table_layout_rows))):
        table_numbers = table_layout_rows[row_idx]
        col_pairs = row_col_pairs[row_idx]
        row_group = row_groups[row_idx]
        for pos_idx, table_no_raw in enumerate(table_numbers):
            if pos_idx >= len(col_pairs):
                break
            try:
                table_no = int(table_no_raw)
            except (TypeError, ValueError):
                continue
            if table_no <= 0 or table_no in slots:
                continue
            left_col, right_col = col_pairs[pos_idx]
            table_slots: list[tuple[int, int]] = []
            for row in row_group:
                table_slots.append((row, left_col))
                table_slots.append((row, right_col))
            slots[table_no] = table_slots
    return slots


def _member_sort_key(row: dict[str, Any]) -> tuple[int, str]:
    try:
        sid = int(row.get("student_id", 0))
    except (TypeError, ValueError):
        sid = 0
    return sid, str(row.get("name", ""))


def _format_member_cell(row: dict[str, Any], include_skill: bool) -> str:
    name = str(row.get("name", "")).strip()
    company = str(row.get("company", "")).strip()
    skill = str(row.get("skill_level", "")).strip()
    parts = [part for part in (name, company) if part]
    if include_skill and skill:
        parts.append(skill)
    return "\n".join(parts)


def _front_to_back_slot(slot: tuple[int, int]) -> tuple[int, int]:
    row, col = slot
    mapped_row = _FRONT_ROW_MAP.get(row, row)
    mapped_col = _FRONT_COL_MAP.get(col, col)
    return mapped_row, mapped_col


def _write_grid_to_worksheet(worksheet: Any, grid: list[list[str]]) -> None:
    for row_idx, row_values in enumerate(grid, start=1):
        for col_idx, value in enumerate(row_values, start=1):
            worksheet.cell(row=row_idx, column=col_idx, value=value)


def _style_excel_worksheet(worksheet: Any, mirrored: bool) -> None:
    _set_excel_dimensions(worksheet)
    worksheet.merge_cells(_TITLE_MERGE_RANGE)
    screen_merge_range = _SCREEN_BOTTOM_MERGE_RANGE if mirrored else _SCREEN_TOP_MERGE_RANGE
    worksheet.merge_cells(screen_merge_range)

    _style_range(
        worksheet,
        _TITLE_MERGE_RANGE,
        fill=_TITLE_FILL,
        border=_THIN_BORDER,
        alignment=_CENTER,
        font=_TITLE_FONT,
    )
    _style_range(
        worksheet,
        screen_merge_range,
        fill=_SCREEN_FILL,
        border=_THIN_BORDER,
        alignment=_CENTER,
        font=_SCREEN_FONT,
    )

    for row, col in _SEAT_COORDS:
        cell = worksheet.cell(row=row, column=col)
        cell.fill = _SEAT_FILL
        cell.border = _THIN_BORDER
        cell.alignment = _CENTER_WRAP
        cell.font = _SEAT_FONT


def _set_excel_dimensions(worksheet: Any) -> None:
    for col in range(1, TEMPLATE_COLS + 1):
        label = _to_col_label(col)
        worksheet.column_dimensions[label].width = 17.0 if col in _SEAT_COLS else 2.8

    for row in range(1, TEMPLATE_ROWS + 1):
        worksheet.row_dimensions[row].height = 20.0
    for row in _SEAT_ROWS:
        worksheet.row_dimensions[row].height = 50.0

    worksheet.row_dimensions[2].height = 28.0
    worksheet.row_dimensions[5].height = 24.0
    worksheet.row_dimensions[19].height = 24.0


def _style_range(
    worksheet: Any,
    cell_range: str,
    fill: PatternFill,
    border: Border,
    alignment: Alignment,
    font: Font,
) -> None:
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = worksheet.cell(row=row, column=col)
            cell.fill = fill
            cell.border = border
            cell.alignment = alignment
            cell.font = font


def _sanitize_title(value: str, max_len: int, invalid_re: re.Pattern[str]) -> str:
    cleaned = invalid_re.sub(" ", value.strip())
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if not cleaned:
        cleaned = "座席表"
    return cleaned[:max_len]


def _unique_title(base: str, used: set[str], max_len: int) -> str:
    candidate = base[:max_len]
    if candidate not in used:
        used.add(candidate)
        return candidate

    counter = 2
    while True:
        suffix = f" ({counter})"
        trimmed = base[: max_len - len(suffix)]
        candidate = f"{trimmed}{suffix}"
        if candidate not in used:
            used.add(candidate)
            return candidate
        counter += 1


def _to_col_label(col_num: int) -> str:
    label = ""
    n = col_num
    while n > 0:
        n, rem = divmod(n - 1, 26)
        label = chr(65 + rem) + label
    return label


def _resolve_template_sheet_name(requested_name: str, available_titles: list[str]) -> str:
    if requested_name in available_titles:
        return requested_name

    normalized_to_original = {_normalize_sheet_title(title): title for title in available_titles}
    normalized_requested = _normalize_sheet_title(requested_name)
    if normalized_requested in normalized_to_original:
        return normalized_to_original[normalized_requested]

    trimmed_date_prefix = re.sub(r"^\d{1,2}月\d{1,2}日", "", requested_name).strip()
    if trimmed_date_prefix in available_titles:
        return trimmed_date_prefix
    normalized_trimmed = _normalize_sheet_title(trimmed_date_prefix)
    if normalized_trimmed in normalized_to_original:
        return normalized_to_original[normalized_trimmed]

    if _DEFAULT_TEMPLATE_SHEET_NAME in available_titles:
        return _DEFAULT_TEMPLATE_SHEET_NAME

    template_like = [title for title in available_titles if "テンプレ" in title]
    if len(template_like) == 1:
        return template_like[0]

    return ""


def _normalize_sheet_title(value: str) -> str:
    return value.replace(" ", "").replace("　", "").strip()
