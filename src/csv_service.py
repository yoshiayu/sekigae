from __future__ import annotations

import csv
import io
from typing import Any

from .constants import CSV_COLUMNS, SKILL_DISPLAY_ORDER, SKILL_LEVELS, skill_to_display


def _normalize_key(value: str) -> str:
    return (
        value.strip()
        .replace("\ufeff", "")
        .replace(" ", "")
        .replace("　", "")
        .replace("_", "")
        .replace("-", "")
        .lower()
    )


HEADER_ALIASES = {
    "name": ["name", "氏名", "名前", "受講生", "受講生名", "受講者", "受講者名"],
    "name_kana": [
        "name_kana",
        "namekana",
        "kana",
        "かな",
        "カナ",
        "furigana",
        "氏名ふりがな",
        "氏名フリガナ",
        "氏名かな",
        "氏名カナ",
        "ふりがな",
        "フリガナ",
        "よみ",
        "読み",
        "よみがな",
        "ヨミガナ",
    ],
    "company": ["company", "会社", "会社名", "企業", "企業名", "所属", "所属会社"],
    "skill_level": [
        "skill_level",
        "skilllevel",
        "skill",
        "スキル",
        "スキルレベル",
        "段階評価",
        "段階",
        "評価",
        "レベル",
    ],
}

HEADER_ALIASES_NORM = {
    key: {_normalize_key(alias) for alias in aliases}
    for key, aliases in HEADER_ALIASES.items()
}


SKILL_ALIASES = {
    "高い": ["高い", "高", "high", "a", "上級"],
    "並": ["並", "普通", "中", "normal", "b", "標準"],
    "低い": ["低い", "やや低", "やや低い", "low", "c", "初級"],
    "ヤバい": ["ヤバい", "ヤバイ", "やばい", "低", "危険", "d", "要支援"],
}

SKILL_ALIASES_NORM: dict[str, str] = {}
for canonical, aliases in SKILL_ALIASES.items():
    for alias in aliases:
        SKILL_ALIASES_NORM[_normalize_key(alias)] = canonical


def template_csv_text() -> str:
    rows = [
        "name,name_kana,company,skill_level",
        "山田太郎,やまだたろう,株式会社A,高",
        "佐藤花子,さとうはなこ,株式会社B,中",
        "鈴木一郎,すずきいちろう,株式会社C,やや低",
    ]
    return "\n".join(rows) + "\n"


def _decode_csv_bytes(file_bytes: bytes) -> str:
    for encoding in ("utf-8-sig", "cp932", "utf-8"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("utf-8", b"", 0, 1, "UTF-8またはCP932でデコードできません。")


def _detect_dialect(text: str) -> csv.Dialect:
    sample = "\n".join(text.splitlines()[:10])
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t")
    except csv.Error:
        return csv.get_dialect("excel")


def _resolve_columns(fieldnames: list[str]) -> tuple[dict[str, str | None], list[str]]:
    column_map: dict[str, str | None] = {
        "name": None,
        "name_kana": None,
        "company": None,
        "skill_level": None,
    }
    warnings: list[str] = []

    for key in column_map.keys():
        aliases_norm = HEADER_ALIASES_NORM[key]
        for field in fieldnames:
            if _normalize_key(field) in aliases_norm:
                column_map[key] = field
                if field.strip() != key:
                    warnings.append(f"列 `{field}` を `{key}` として読み替えました。")
                break

    return column_map, warnings


def _normalize_skill_level(value: str) -> str | None:
    raw = value.strip()
    if raw in SKILL_LEVELS:
        return raw
    return SKILL_ALIASES_NORM.get(_normalize_key(raw))


def parse_students_csv(file_bytes: bytes) -> tuple[list[dict[str, str]], list[str]]:
    text = _decode_csv_bytes(file_bytes)
    stream = io.StringIO(text)
    reader = csv.DictReader(stream, dialect=_detect_dialect(text))

    if reader.fieldnames is None:
        return [], ["ヘッダー行が見つかりません。`name,company,skill_level` を含めてください。"]

    fieldnames = [h.strip() for h in reader.fieldnames if h is not None]
    column_map, warnings = _resolve_columns(fieldnames)
    missing = [col for col in ("name", "company") if column_map[col] is None]
    if missing:
        expected = ", ".join(CSV_COLUMNS)
        found = ", ".join(fieldnames) if fieldnames else "(none)"
        return [], [
            "ヘッダー不正: 必須列が不足しています。"
            f" missing={','.join(missing)} / expected={expected} / found={found}"
        ]

    rows: list[dict[str, str]] = []
    errors: list[str] = []
    defaulted_skill_count = 0

    for line_no, row in enumerate(reader, start=2):
        name = (row.get(str(column_map["name"])) or "").strip()
        name_kana = ""
        if column_map["name_kana"] is not None:
            name_kana = (row.get(str(column_map["name_kana"])) or "").strip()
        company = (row.get(str(column_map["company"])) or "").strip()
        skill_raw = ""
        if column_map["skill_level"] is not None:
            skill_raw = (row.get(str(column_map["skill_level"])) or "").strip()

        if not name and not company and not skill_raw:
            continue

        row_errors: list[str] = []
        if not name:
            row_errors.append("name が空です")
        if not company:
            row_errors.append("company が空です")

        if skill_raw:
            skill_level = _normalize_skill_level(skill_raw)
            if skill_level is None:
                allowed = " / ".join(SKILL_DISPLAY_ORDER)
                row_errors.append(f"skill_level が不正です（{allowed} のみ可）")
            else:
                normalized_skill = skill_level
        else:
            normalized_skill = "並"
            defaulted_skill_count += 1

        if row_errors:
            errors.append(f"{line_no}行目: " + ", ".join(row_errors))
            continue

        rows.append(
            {
                "name": name,
                "name_kana": name_kana,
                "company": company,
                "skill_level": normalized_skill,
            }
        )

    # 段階評価が無い/空欄でも実務運用しやすいよう、既定値「並」で取り込みます。
    if column_map["skill_level"] is None or defaulted_skill_count > 0:
        _ = warnings

    return rows, errors


def rows_to_assignment_csv(
    rows: list[dict[str, Any]], target_week: str = "", history_id: int | None = None
) -> bytes:
    output = io.StringIO()
    writer = csv.writer(output, lineterminator="\n")
    writer.writerow(
        [
            "table_no",
            "student_id",
            "name",
            "name_kana",
            "company",
            "skill_level",
            "target_week",
            "seating_history_id",
        ]
    )
    for row in rows:
        writer.writerow(
            [
                row["table_no"],
                row["student_id"],
                row["name"],
                row.get("name_kana", ""),
                row["company"],
                skill_to_display(str(row["skill_level"])),
                target_week,
                history_id if history_id is not None else "",
            ]
        )
    return output.getvalue().encode("utf-8-sig")
