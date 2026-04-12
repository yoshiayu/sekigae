from __future__ import annotations

import json
import re
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from src.constants import (
    DEFAULT_ATTEMPTS,
    DEFAULT_MAX_PER_TABLE,
    DEFAULT_MIN_PER_TABLE,
    DEFAULT_TABLE_COUNT,
    DEFAULT_WEIGHTS,
    SKILL_COLORS,
    SKILL_DISPLAY_ORDER,
    SKILL_LEVELS,
    skill_to_display,
)
from src.csv_service import parse_students_csv, rows_to_assignment_csv, template_csv_text
from src.db import init_db
from src.layout_export import (
    build_layout_payload,
    default_daily_sheet_name,
    default_table_layout_rows,
    export_layouts_to_excel_bytes,
    publish_layouts_to_google_sheets,
    safe_download_filename,
)
from src.repository import (
    bulk_insert_students,
    bulk_update_student_skills_by_name,
    create_seating_history,
    create_student,
    delete_all_histories,
    delete_all_students,
    delete_student,
    get_history,
    get_history_rows,
    get_previous_context,
    list_histories,
    list_students,
    save_assignments,
    update_student,
)
from src.seating import (
    SeatingConfig,
    Student,
    build_table_view,
    count_soft_avoid_pairs,
    evaluate_rows,
    generate_best_assignment,
    optimize_soft_skill_conflicts,
    validate_manual_rows,
)


st.set_page_config(page_title="受講生席替えアプリ", page_icon="席", layout="wide")
init_db()

_LAYOUT_ROW_LABELS = ("前列（スクリーン側）", "中列", "後列（後方）")
_LAYOUT_ROW_LIMITS = (4, 5, 4)


def _current_iso_week() -> str:
    return datetime.now().strftime("%Y-W%V")


def _skill_distribution_text(distribution: dict[str, int]) -> str:
    total = sum(distribution.values())
    if total == 0:
        return "データなし"
    chunks = []
    for skill in SKILL_LEVELS:
        count = distribution.get(skill, 0)
        ratio = (count / total) * 100
        chunks.append(f"{skill_to_display(skill)}: {count}名 ({ratio:.1f}%)")
    return " / ".join(chunks)


def _rows_to_students(rows: list[dict[str, Any]]) -> list[Student]:
    return [
        Student(
            id=int(row["id"]),
            name=str(row["name"]),
            company=str(row["company"]),
            skill_level=str(row["skill_level"]),
            name_kana=str(row.get("name_kana", "")),
        )
        for row in rows
    ]


def _skill_badge(skill: str) -> str:
    color = SKILL_COLORS.get(skill, "#4A5568")
    label = skill_to_display(skill)
    return (
        f"<span style='background:{color};color:#fff;border-radius:999px;"
        f"padding:2px 8px;font-size:12px;'>{label}</span>"
    )


def _config_from_settings(settings: dict[str, Any] | None) -> SeatingConfig:
    if not isinstance(settings, dict):
        return SeatingConfig()
    allowed = set(SeatingConfig.__dataclass_fields__.keys())
    filtered = {key: value for key, value in settings.items() if key in allowed}
    return SeatingConfig(**filtered)


def _layout_row_to_text(row: list[int]) -> str:
    return ",".join(str(no) for no in row)


def _default_layout_rows_for_count(table_count: int) -> list[list[int]]:
    return [list(row) for row in default_table_layout_rows(max(0, int(table_count)))]


def _parse_layout_row_text(value: str) -> tuple[list[int], list[str]]:
    tokens = [token for token in re.split(r"[,\s、]+", str(value).strip()) if token]
    row: list[int] = []
    invalid: list[str] = []
    for token in tokens:
        try:
            row.append(int(token))
        except ValueError:
            invalid.append(token)
    return row, invalid


def _validate_layout_rows(layout_rows: list[list[int]], table_count: int) -> list[str]:
    errors: list[str] = []
    if len(layout_rows) != 3:
        return ["レイアウト行数が不正です。前列・中列・後列の3行を設定してください。"]

    for idx, (row, limit) in enumerate(zip(layout_rows, _LAYOUT_ROW_LIMITS)):
        if len(row) > limit:
            errors.append(
                f"{_LAYOUT_ROW_LABELS[idx]}は最大{limit}テーブルまでです。"
            )

    flat = [no for row in layout_rows for no in row]
    out_of_range = sorted({no for no in flat if no < 1 or no > table_count})
    if out_of_range:
        errors.append(
            f"レイアウトに範囲外のテーブル番号があります: {', '.join(str(no) for no in out_of_range)}"
        )

    duplicates = sorted({no for no in flat if flat.count(no) > 1})
    if duplicates:
        errors.append(
            f"レイアウトに重複テーブル番号があります: {', '.join(str(no) for no in duplicates)}"
        )

    expected = set(range(1, table_count + 1))
    provided = set(flat)
    missing = sorted(expected - provided)
    if missing:
        errors.append(
            f"レイアウト未指定のテーブルがあります: {', '.join(str(no) for no in missing)}"
        )
    return errors


def _coerce_layout_rows(
    candidate: Any,
    table_count: int,
) -> list[list[int]]:
    default_rows = _default_layout_rows_for_count(table_count)
    if not isinstance(candidate, list) or len(candidate) != 3:
        return default_rows
    try:
        rows = [[int(v) for v in row] for row in candidate]
    except (TypeError, ValueError):
        return default_rows
    if _validate_layout_rows(rows, table_count):
        return default_rows
    return rows


def _ensure_layout_editor_state(table_count: int, layout_rows: list[list[int]]) -> None:
    marker_key = "layout_editor_table_count"
    marker_sig_key = "layout_editor_signature"
    signature = "|".join(_layout_row_to_text(row) for row in layout_rows)
    if (
        st.session_state.get(marker_key) != table_count
        or st.session_state.get(marker_sig_key) != signature
    ):
        st.session_state["layout_front_text"] = _layout_row_to_text(layout_rows[0])
        st.session_state["layout_middle_text"] = _layout_row_to_text(layout_rows[1])
        st.session_state["layout_back_text"] = _layout_row_to_text(layout_rows[2])
        st.session_state[marker_key] = table_count
        st.session_state[marker_sig_key] = signature


def _read_layout_rows_from_editor(table_count: int) -> tuple[list[list[int]], list[str]]:
    row_texts = [
        st.session_state.get("layout_front_text", ""),
        st.session_state.get("layout_middle_text", ""),
        st.session_state.get("layout_back_text", ""),
    ]
    parsed_rows: list[list[int]] = []
    errors: list[str] = []
    for idx, text in enumerate(row_texts):
        row, invalid = _parse_layout_row_text(text)
        parsed_rows.append(row)
        if invalid:
            errors.append(
                f"{_LAYOUT_ROW_LABELS[idx]}に整数以外の値があります: {', '.join(invalid)}"
            )
    errors.extend(_validate_layout_rows(parsed_rows, table_count))
    return parsed_rows, errors


def _make_config_from_state() -> SeatingConfig:
    settings = st.session_state.get("current_settings")
    return _config_from_settings(settings if isinstance(settings, dict) else None)


def _store_result(
    rows: list[dict[str, Any]],
    score: float,
    metrics: dict[str, Any],
    config: SeatingConfig,
    target_week: str,
    table_layout_rows: list[list[int]],
    previous_history_id: int | None,
    previous_pairs: set[tuple[int, int]],
    previous_table_map: dict[int, int],
) -> None:
    normalized_layout_rows = _coerce_layout_rows(table_layout_rows, config.table_count)
    st.session_state["current_rows"] = rows
    st.session_state["current_score"] = score
    st.session_state["current_metrics"] = metrics
    st.session_state["current_target_week"] = target_week
    st.session_state["current_layout_rows"] = normalized_layout_rows
    st.session_state["current_settings"] = {
        "table_count": config.table_count,
        "min_per_table": config.min_per_table,
        "max_per_table": config.max_per_table,
        "attempts": config.attempts,
        "company_weight": config.company_weight,
        "previous_weight": config.previous_weight,
        "skill_weight": config.skill_weight,
        "size_weight": config.size_weight,
        "randomness_weight": config.randomness_weight,
        "seed": config.seed,
        "table_layout_rows": normalized_layout_rows,
    }
    st.session_state["previous_history_id"] = previous_history_id
    st.session_state["previous_pairs"] = previous_pairs
    st.session_state["previous_table_map"] = previous_table_map


def _clear_current_result_state() -> None:
    keys = [
        "current_rows",
        "current_score",
        "current_metrics",
        "current_target_week",
        "current_layout_rows",
        "current_settings",
        "previous_history_id",
        "previous_pairs",
        "previous_table_map",
    ]
    for key in keys:
        st.session_state.pop(key, None)


def _load_dummy_csv() -> tuple[list[dict[str, str]], list[str]]:
    dummy_path = Path("data/dummy_students_73.csv")
    if not dummy_path.exists():
        return [], ["data/dummy_students_73.csv が見つかりません。"]
    return parse_students_csv(dummy_path.read_bytes())


def _normalize_name_text(value: str) -> str:
    return " ".join(value.strip().replace("　", " ").split())


def _normalize_skill_input(value: str) -> str | None:
    raw = value.strip()
    if raw in SKILL_LEVELS:
        return raw

    compact = _normalize_name_text(raw).replace(" ", "").lower()
    aliases: dict[str, str] = {
        "高い": "高い",
        "高": "高い",
        "high": "高い",
        "a": "高い",
        "上級": "高い",
        "並": "並",
        "普通": "並",
        "中": "並",
        "normal": "並",
        "b": "並",
        "標準": "並",
        "低い": "低い",
        "やや低": "低い",
        "やや低い": "低い",
        "low": "低い",
        "c": "低い",
        "初級": "低い",
        "ヤバい": "ヤバい",
        "ヤバイ": "ヤバい",
        "やばい": "ヤバい",
        "低": "ヤバい",
        "危険": "ヤバい",
        "d": "ヤバい",
        "要支援": "ヤバい",
    }
    return aliases.get(compact)


def _split_name_skill_line(line: str) -> tuple[str, str] | None:
    if "\t" in line:
        tab_cells = [cell.strip() for cell in line.split("\t")]
        non_empty = [cell for cell in tab_cells if cell]
        if len(non_empty) >= 2:
            return non_empty[0], non_empty[1]

    for separator in ("=", "＝", ":", "：", ",", "，", "、"):
        if separator in line:
            left, right = line.split(separator, 1)
            return left.strip(), right.strip()
    return None


def _is_skill_header(name_text: str, skill_text: str) -> bool:
    normalized_name = _normalize_name_text(name_text).replace(" ", "").lower()
    normalized_skill = _normalize_name_text(skill_text).replace(" ", "").lower()
    name_candidates = {"受講生", "受講生名", "受講者", "受講者名", "氏名", "名前", "name"}
    skill_candidates = {
        "スキル",
        "スキルレベル",
        "段階評価",
        "評価",
        "レベル",
        "skill",
        "skilllevel",
        "skill_level",
    }
    return normalized_name in name_candidates and normalized_skill in skill_candidates


def _parse_skill_updates_text(text: str) -> tuple[list[tuple[str, str]], list[str]]:
    rows: list[tuple[str, str]] = []
    errors: list[str] = []

    for line_no, raw_line in enumerate(text.splitlines(), start=1):
        line = raw_line.strip()
        if not line:
            continue

        split_result = _split_name_skill_line(line)
        if split_result is None:
            errors.append(f"{line_no}行目: 区切り文字（タブ / = / : / ,）が見つかりません。")
            continue

        name_raw, skill_raw = split_result
        if _is_skill_header(name_raw, skill_raw) and not rows:
            continue

        name = _normalize_name_text(name_raw)
        if not name:
            errors.append(f"{line_no}行目: 受講生名が空です。")
            continue

        normalized_skill = _normalize_skill_input(skill_raw)
        if normalized_skill is None:
            allowed = " / ".join(SKILL_DISPLAY_ORDER)
            errors.append(f"{line_no}行目: スキルが不正です（{allowed} のみ可）。")
            continue

        rows.append((name, normalized_skill))

    if not rows and not errors:
        errors.append("貼り付けデータが空です。")

    return rows, errors


st.title("受講生席替えアプリ（MVP）")
st.caption("毎週運用を想定した、スキル分散 + 会社分散 + 前回同席回避 の自動席替え")

tab_students, tab_run, tab_result, tab_history = st.tabs(
    ["受講生管理", "席替え実行", "席替え結果", "履歴"]
)


with tab_students:
    st.subheader("受講生一覧")
    flash = st.session_state.pop("bulk_skill_update_flash", None)
    if isinstance(flash, dict):
        requested_names = int(flash.get("requested_names", 0))
        matched_names = int(flash.get("matched_names", 0))
        updated_rows = int(flash.get("updated_rows", 0))
        if updated_rows > 0:
            st.success(
                f"{requested_names}件の指定を受け取り、{matched_names}名に一致・"
                f"{updated_rows}件のスキルを更新しました。"
            )
        else:
            st.warning("一致する受講生が見つからず、更新は0件でした。")

        unmatched_names = [str(name) for name in flash.get("unmatched_names", [])]
        if unmatched_names:
            preview = " / ".join(unmatched_names[:10])
            suffix = " ..." if len(unmatched_names) > 10 else ""
            st.info(f"未一致の受講生: {preview}{suffix}")

        overwritten_names = [str(name) for name in flash.get("overwritten_names", [])]
        if overwritten_names:
            preview = " / ".join(overwritten_names[:10])
            suffix = " ..." if len(overwritten_names) > 10 else ""
            st.info(f"同名の複数指定は最後の行を採用: {preview}{suffix}")

    search = st.text_input("検索（氏名 / 会社名）", value="", key="student_search")
    students = list_students(search=search)
    all_students = list_students()
    distribution = Counter(row["skill_level"] for row in all_students)

    left, right = st.columns([2, 1])
    with left:
        st.write(f"登録人数: **{len(all_students)}名**")
        st.write(_skill_distribution_text(distribution))
    with right:
        if st.button("ダミー73名を全置換で投入", key="load_dummy_students"):
            parsed_rows, errs = _load_dummy_csv()
            if errs:
                for err in errs:
                    st.error(err)
            else:
                delete_all_students()
                bulk_insert_students(parsed_rows)
                st.success("ダミーデータ73名を登録しました。")
                st.rerun()

        st.divider()
        reset_confirm = st.checkbox(
            "受講生・席替え履歴・画面結果を全リセットしてよい",
            key="reset_all_confirm",
        )
        if st.button(
            "全リセット実行",
            key="reset_all_button",
            type="secondary",
            disabled=not reset_confirm,
        ):
            delete_all_histories()
            delete_all_students()
            _clear_current_result_state()
            st.success("全リセットしました。")
            st.rerun()

    if students:
        df_students = pd.DataFrame(students)
        df_students_display = df_students.copy()
        if "skill_level" in df_students_display.columns:
            df_students_display["skill_level"] = df_students_display["skill_level"].map(
                skill_to_display
            )
        st.dataframe(
            df_students_display[
                ["id", "name_kana", "name", "company", "skill_level", "updated_at"]
            ],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("該当する受講生がいません。")

    st.divider()
    col_add, col_edit = st.columns(2)

    with col_add:
        st.markdown("#### 受講生追加")
        with st.form("add_student_form", clear_on_submit=True):
            add_name = st.text_input("氏名", key="add_name")
            add_name_kana = st.text_input("氏名ふりがな（任意）", key="add_name_kana")
            add_company = st.text_input("会社名", key="add_company")
            add_skill = st.selectbox(
                "スキル",
                options=list(SKILL_LEVELS),
                format_func=skill_to_display,
                key="add_skill",
            )
            add_submit = st.form_submit_button("追加")
        if add_submit:
            if not add_name.strip() or not add_company.strip():
                st.error("氏名・会社名は必須です。")
            else:
                create_student(add_name, add_company, add_skill, name_kana=add_name_kana)
                st.success("受講生を追加しました。")
                st.rerun()

    with col_edit:
        st.markdown("#### 受講生編集 / 削除")
        all_students_for_edit = list_students()
        if all_students_for_edit:
            selected_id = st.selectbox(
                "編集対象",
                options=[int(row["id"]) for row in all_students_for_edit],
                format_func=lambda sid: next(
                    (
                        f"{row['id']}: {row['name']} ({row['company']})"
                        for row in all_students_for_edit
                        if int(row["id"]) == int(sid)
                    ),
                    str(sid),
                ),
                key="edit_student_id",
            )
            selected = next(
                row for row in all_students_for_edit if int(row["id"]) == int(selected_id)
            )

            # 選択IDごとにフォームキーを分けて、対象変更時に値が追従するようにする
            form_key = f"edit_student_form_{selected_id}"
            with st.form(form_key):
                edit_name = st.text_input(
                    "氏名", value=str(selected["name"]), key=f"edit_name_{selected_id}"
                )
                edit_name_kana = st.text_input(
                    "氏名ふりがな（任意）",
                    value=str(selected.get("name_kana", "")),
                    key=f"edit_name_kana_{selected_id}",
                )
                edit_company = st.text_input(
                    "会社名",
                    value=str(selected["company"]),
                    key=f"edit_company_{selected_id}",
                )
                edit_skill = st.selectbox(
                    "スキル",
                    options=list(SKILL_LEVELS),
                    index=list(SKILL_LEVELS).index(str(selected["skill_level"])),
                    format_func=skill_to_display,
                    key=f"edit_skill_{selected_id}",
                )
                edit_submit = st.form_submit_button("更新")

            if edit_submit:
                if not edit_name.strip() or not edit_company.strip():
                    st.error("氏名・会社名は必須です。")
                else:
                    update_student(
                        int(selected_id),
                        edit_name,
                        edit_company,
                        edit_skill,
                        name_kana=edit_name_kana,
                    )
                    st.success("更新しました。")
                    st.rerun()

            delete_confirm = st.checkbox("削除を実行してよい", key="delete_confirm")
            if st.button(
                "選択した受講生を削除",
                type="secondary",
                disabled=not delete_confirm,
                key="delete_student_button",
            ):
                delete_student(int(selected_id))
                st.success("削除しました。")
                st.rerun()
        else:
            st.info("受講生データがありません。")

    st.divider()
    st.markdown("#### CSVインポート / テンプレート")
    csv_file = st.file_uploader("CSVファイルを選択", type=["csv"], key="students_csv_uploader")
    replace_all = st.checkbox("既存受講生を削除して全置換する", key="replace_all_csv")
    if st.button("CSV取り込み実行", key="import_students_csv"):
        if csv_file is None:
            st.error("CSVファイルを選択してください。")
        else:
            parsed_rows, errors = parse_students_csv(csv_file.getvalue())
            if errors:
                st.error("CSV取り込みでエラーが発生しました。")
                for err in errors:
                    st.write(f"- {err}")
            else:
                if replace_all:
                    delete_all_students()
                inserted = bulk_insert_students(parsed_rows)
                skipped = len(parsed_rows) - inserted
                st.success(f"{inserted}件を取り込みました。")
                if skipped > 0:
                    st.info(f"{skipped}件は既存データと重複のためスキップしました。")
                st.rerun()

    st.divider()
    st.markdown("#### スキル一括更新（名前は変更しません）")
    st.caption(
        "Excelの2列（受講生 / スキル）をそのまま貼り付けできます。"
        " 区切りはタブ・`=`・`:`・`,` に対応。"
    )
    bulk_skill_text = st.text_area(
        "貼り付けデータ",
        key="bulk_skill_text",
        height=180,
        placeholder="受講生\tスキル\n山田 太郎\t高\n佐藤 花子\t中",
    )
    if st.button("スキル一括更新を実行", key="bulk_update_skill_button"):
        parsed_updates, parse_errors = _parse_skill_updates_text(bulk_skill_text)
        if parse_errors:
            st.error("一括更新データに不備があります。")
            for err in parse_errors:
                st.write(f"- {err}")
        else:
            name_to_skill: dict[str, str] = {}
            overwritten_names: list[str] = []
            for name, skill_level in parsed_updates:
                previous_skill = name_to_skill.get(name)
                if previous_skill is not None and previous_skill != skill_level:
                    overwritten_names.append(name)
                name_to_skill[name] = skill_level

            updated_rows, unmatched_names, matched_name_counts = bulk_update_student_skills_by_name(
                name_to_skill
            )
            st.session_state["bulk_skill_update_flash"] = {
                "requested_names": len(name_to_skill),
                "matched_names": len(matched_name_counts),
                "updated_rows": updated_rows,
                "unmatched_names": unmatched_names,
                "overwritten_names": list(dict.fromkeys(overwritten_names)),
            }
            st.rerun()

    st.download_button(
        label="CSVテンプレートをダウンロード",
        data=template_csv_text().encode("utf-8-sig"),
        file_name="students_template.csv",
        mime="text/csv",
        key="download_template_csv",
    )


with tab_run:
    st.subheader("席替え実行")
    students_rows = list_students()
    student_count = len(students_rows)
    distribution = Counter(row["skill_level"] for row in students_rows)
    col_table, col_min, col_max = st.columns(3)
    with col_table:
        table_count = int(
            st.number_input(
                "使用テーブル数",
                min_value=1,
                max_value=DEFAULT_TABLE_COUNT,
                value=DEFAULT_TABLE_COUNT,
                step=1,
                key="run_table_count",
            )
        )
    with col_min:
        min_per_table = int(
            st.number_input(
                "1テーブル最小人数",
                min_value=1,
                max_value=DEFAULT_MAX_PER_TABLE,
                value=DEFAULT_MIN_PER_TABLE,
                step=1,
                key="run_min_per_table",
            )
        )
    with col_max:
        max_per_table = int(
            st.number_input(
                "1テーブル最大人数",
                min_value=min_per_table,
                max_value=DEFAULT_MAX_PER_TABLE,
                value=max(min_per_table, DEFAULT_MAX_PER_TABLE),
                step=1,
                key="run_max_per_table",
            )
        )

    min_required = table_count * min_per_table
    capacity = table_count * max_per_table

    col_info1, col_info2, col_info3, col_info4 = st.columns(4)
    col_info1.metric("登録人数", f"{student_count}名")
    col_info2.metric("最小必要人数", f"{min_required}名")
    col_info3.metric("最大収容人数", f"{capacity}名")
    if student_count < min_required:
        col_info4.metric("不足", f"{min_required - student_count}名")
    elif student_count > capacity:
        col_info4.metric("超過", f"{student_count - capacity}名")
    else:
        col_info4.metric("空席", f"{capacity - student_count}席")
    st.write(_skill_distribution_text(distribution))

    if student_count < min_required:
        st.error(
            f"受講生数 {student_count} 名は最小必要人数 {min_required} 名を下回っています。"
            " テーブル数を減らすか、最小人数を下げてください。"
        )
    if student_count > capacity:
        st.error(
            f"受講生数 {student_count} 名は上限 {capacity} 名を超えています。"
            " テーブル数か最大人数を増やす必要があります。"
        )

    st.markdown("#### 条件設定")
    target_week = st.text_input("対象週", value=_current_iso_week(), key="target_week")
    col_w1, col_w2, col_w3 = st.columns(3)
    with col_w1:
        company_weight = st.slider(
            "同じ会社回避の強さ", 1.0, 60.0, float(DEFAULT_WEIGHTS["company"]), 1.0
        )
        previous_weight = st.slider(
            "前回同席回避の強さ", 1.0, 40.0, float(DEFAULT_WEIGHTS["previous"]), 1.0
        )
    with col_w2:
        skill_weight = st.slider(
            "同スキル固めの強さ（強化）",
            1.0,
            60.0,
            float(DEFAULT_WEIGHTS["skill"]),
            1.0,
        )
        size_weight = st.slider("人数均等化の強さ", 0.0, 10.0, float(DEFAULT_WEIGHTS["size"]), 0.5)
    st.caption(
        "スキル同席制約: 高×やや低 / 高×低 は同席しません。"
        " 中×低 は可能な限り回避し、"
        " 高→中→やや低→低 の隣接スキル同士を優先して混在させます。"
    )
    with col_w3:
        randomness_weight = st.slider(
            "ランダム性", 0.0, 5.0, float(DEFAULT_WEIGHTS["randomness"]), 0.1
        )
        attempts = st.slider("試行回数", 100, 500, DEFAULT_ATTEMPTS, 50)
        seed_text = st.text_input("乱数シード（任意）", value="", key="seed_text")

    run_disabled = student_count == 0 or student_count < min_required or student_count > capacity
    if st.button("席替えを実行", type="primary", disabled=run_disabled, key="run_seating"):
        try:
            seed_value = int(seed_text) if seed_text.strip() else None
        except ValueError:
            st.error("乱数シードは整数で入力してください。")
            seed_value = None
        else:
            previous_history_id, previous_pairs, previous_table_map = get_previous_context()
            config = SeatingConfig(
                table_count=table_count,
                min_per_table=min_per_table,
                max_per_table=max_per_table,
                attempts=attempts,
                company_weight=company_weight,
                previous_weight=previous_weight,
                skill_weight=skill_weight,
                size_weight=size_weight,
                randomness_weight=randomness_weight,
                seed=seed_value,
            )
            try:
                result = generate_best_assignment(
                    students=_rows_to_students(students_rows),
                    previous_pairs=previous_pairs,
                    previous_table_map=previous_table_map,
                    config=config,
                )
            except (ValueError, RuntimeError) as exc:
                st.error(str(exc))
                result = None
            if result is not None:
                run_layout_rows = _coerce_layout_rows(
                    st.session_state.get("current_layout_rows"),
                    config.table_count,
                )
                _store_result(
                    rows=result["rows"],
                    score=float(result["score"]),
                    metrics=result["metrics"],
                    config=config,
                    target_week=target_week,
                    table_layout_rows=run_layout_rows,
                    previous_history_id=previous_history_id,
                    previous_pairs=previous_pairs,
                    previous_table_map=previous_table_map,
                )
                st.success("席替えを生成しました。結果タブで確認できます。")


with tab_result:
    st.subheader("席替え結果")
    current_rows: list[dict[str, Any]] = st.session_state.get("current_rows", [])
    if not current_rows:
        st.info("先に「席替え実行」タブで席替えを実行してください。")
    else:
        current_config = _make_config_from_state()
        previous_pairs = st.session_state.get("previous_pairs", set())
        previous_table_map = st.session_state.get("previous_table_map", {})
        current_metrics = st.session_state.get("current_metrics", {})
        current_score = float(st.session_state.get("current_score", 0.0))
        current_target_week = st.session_state.get("current_target_week", _current_iso_week())

        metric_cols = st.columns(4)
        metric_cols[0].metric("総合スコア", f"{current_score:.1f}")
        metric_cols[1].metric(
            "前回同席ペア重複率",
            f"{100 * float(current_metrics.get('overlap_pair_rate', 0.0)):.1f}%",
        )
        metric_cols[2].metric(
            "前回と同じテーブル率",
            f"{100 * float(current_metrics.get('same_table_student_rate', 0.0)):.1f}%",
        )
        metric_cols[3].metric("会社重複数", int(current_metrics.get("company_collision_total", 0)))
        forbidden_pairs = int(current_metrics.get("forbidden_skill_pair_total", 0))
        if forbidden_pairs > 0:
            st.warning(
                f"収容優先で配置したため、禁止スキル同席が {forbidden_pairs} 組あります。"
                "（高×やや低 / 高×低）"
            )
        elif bool(current_metrics.get("fallback_used", False)):
            st.info("厳密解が見つからなかったため、収容優先モードで席替えしています。")
        soft_avoid_pairs = int(
            current_metrics.get(
                "soft_avoid_pair_total",
                count_soft_avoid_pairs(current_rows, current_config.table_count),
            )
        )
        if soft_avoid_pairs > 0:
            st.warning(
                f"中×低 の同席ペアが {soft_avoid_pairs} 組あります。"
                "下の自動解消ボタンで改善できます。"
            )
            if st.button("中×低同席を自動解消", key="auto_fix_soft_conflicts"):
                fixed_rows = optimize_soft_skill_conflicts(
                    rows=current_rows,
                    table_count=current_config.table_count,
                )
                before_soft = count_soft_avoid_pairs(current_rows, current_config.table_count)
                after_soft = count_soft_avoid_pairs(fixed_rows, current_config.table_count)
                if after_soft < before_soft:
                    fixed_score, fixed_metrics = evaluate_rows(
                        rows=fixed_rows,
                        previous_pairs=previous_pairs,
                        previous_table_map=previous_table_map,
                        config=current_config,
                    )
                    st.session_state["current_rows"] = fixed_rows
                    st.session_state["current_score"] = fixed_score
                    st.session_state["current_metrics"] = fixed_metrics
                    st.success(
                        f"中×低同席ペアを {before_soft} → {after_soft} に改善しました。"
                    )
                    st.rerun()
                else:
                    st.info("自動解消を試しましたが、これ以上は改善できませんでした。")

        st.markdown("#### 手動調整（table_no を直接変更）")
        df_current = pd.DataFrame(current_rows)
        skill_by_student_id = {
            int(row["student_id"]): str(row["skill_level"]) for row in current_rows
        }
        df_current_display = df_current.copy()
        df_current_display["skill_label"] = df_current_display["skill_level"].map(
            skill_to_display
        )
        edited_df = st.data_editor(
            df_current_display[
                ["table_no", "student_id", "name_kana", "name", "company", "skill_label"]
            ],
            hide_index=True,
            use_container_width=True,
            disabled=["student_id", "name_kana", "name", "company", "skill_label"],
            column_config={
                "skill_label": "スキル",
                "table_no": st.column_config.NumberColumn(
                    "table_no", min_value=1, max_value=current_config.table_count, step=1
                ),
            },
            key="result_editor",
        )

        if st.button("手動調整を反映", key="apply_manual_changes"):
            updated_rows = edited_df.to_dict(orient="records")
            for row in updated_rows:
                row["skill_level"] = skill_by_student_id.get(int(row["student_id"]), "")
                row.pop("skill_label", None)
            errors = validate_manual_rows(
                rows=updated_rows,
                table_count=current_config.table_count,
                min_per_table=current_config.min_per_table,
                max_per_table=current_config.max_per_table,
            )
            if errors:
                for err in errors:
                    st.error(err)
            else:
                updated_rows.sort(key=lambda r: (int(r["table_no"]), int(r["student_id"])))
                score, metrics = evaluate_rows(
                    rows=updated_rows,
                    previous_pairs=previous_pairs,
                    previous_table_map=previous_table_map,
                    config=current_config,
                )
                st.session_state["current_rows"] = updated_rows
                st.session_state["current_score"] = score
                st.session_state["current_metrics"] = metrics
                st.success("手動調整を反映しました。")
                st.rerun()

        settings_obj = st.session_state.get("current_settings", {})
        settings_layout_rows = (
            settings_obj.get("table_layout_rows")
            if isinstance(settings_obj, dict)
            else None
        )
        initial_layout_rows = _coerce_layout_rows(
            st.session_state.get("current_layout_rows") or settings_layout_rows,
            current_config.table_count,
        )
        st.session_state["current_layout_rows"] = initial_layout_rows
        _ensure_layout_editor_state(current_config.table_count, initial_layout_rows)

        st.markdown("#### テーブル配置設定（前4・中5・後4の千鳥）")
        layout_col1, layout_col2, layout_col3 = st.columns(3)
        with layout_col1:
            st.text_input(
                "前列（左→右の table_no）",
                key="layout_front_text",
                help="例: 1,4,7,10",
            )
        with layout_col2:
            st.text_input(
                "中列（左→右の table_no）",
                key="layout_middle_text",
                help="例: 2,5,8,11,13",
            )
        with layout_col3:
            st.text_input(
                "後列（左→右の table_no）",
                key="layout_back_text",
                help="例: 3,6,9,12",
            )

        if st.button("配置設定を反映", key="apply_layout_rows"):
            edited_layout_rows, layout_errors = _read_layout_rows_from_editor(
                current_config.table_count
            )
            if layout_errors:
                for err in layout_errors:
                    st.error(err)
            else:
                st.session_state["current_layout_rows"] = edited_layout_rows
                if isinstance(settings_obj, dict):
                    settings_obj["table_layout_rows"] = edited_layout_rows
                    st.session_state["current_settings"] = settings_obj
                st.success("配置設定を反映しました。")
                st.rerun()

        active_layout_rows = _coerce_layout_rows(
            st.session_state.get("current_layout_rows"),
            current_config.table_count,
        )

        st.markdown("#### テーブル別表示")
        table_map = build_table_view(current_rows, current_config.table_count)
        table_metrics = {
            int(m["table_no"]): m for m in st.session_state["current_metrics"].get("table_metrics", [])
        }
        for row_label, row_table_nos in zip(_LAYOUT_ROW_LABELS, active_layout_rows):
            st.markdown(f"##### {row_label}")
            if not row_table_nos:
                st.caption("（設定なし）")
                continue
            row_cols = st.columns(len(row_table_nos))
            for idx, table_no in enumerate(row_table_nos):
                members = table_map.get(table_no, [])
                metric = table_metrics.get(table_no, {})
                with row_cols[idx]:
                    with st.container(border=True):
                        st.markdown(f"##### Table {table_no} ({len(members)}名)")
                        if not members:
                            st.caption("未配置")
                        for member in members:
                            badge = _skill_badge(str(member["skill_level"]))
                            kana = str(member.get("name_kana", "")).strip()
                            kana_html = f"{kana}<br>" if kana else ""
                            st.markdown(
                                (
                                    f"<div style='padding:6px 8px;border:1px solid #E2E8F0;"
                                    f"border-radius:8px;margin-bottom:6px;'>"
                                    f"{kana_html}"
                                    f"<strong>{member['name']}</strong><br>"
                                    f"{member['company']}<br>{badge}</div>"
                                ),
                                unsafe_allow_html=True,
                            )

                        duplicate_companies = metric.get("duplicate_companies", {})
                        if duplicate_companies:
                            info = ", ".join(f"{k}({v})" for k, v in duplicate_companies.items())
                            st.warning(f"会社重複: {info}")

                        prev_hits = int(metric.get("previous_pair_hits", 0))
                        if prev_hits > 0:
                            st.caption(f"前回同席ペア重複: {prev_hits}")

        displayed_tables = {no for row in active_layout_rows for no in row}
        hidden_tables = [
            table_no
            for table_no in range(1, current_config.table_count + 1)
            if table_no not in displayed_tables
        ]
        if hidden_tables:
            st.warning(
                "配置設定に含まれていないため非表示のテーブルがあります: "
                + ", ".join(str(no) for no in hidden_tables)
            )

        csv_bytes = rows_to_assignment_csv(
            rows=current_rows,
            target_week=str(current_target_week),
            history_id=None,
        )
        st.download_button(
            label="現在の結果をCSV出力",
            data=csv_bytes,
            file_name=f"seating_result_{current_target_week}.csv",
            mime="text/csv",
            key="download_current_result_csv",
        )

        st.markdown("#### スプレッドシート形式出力（Excel / Googleスプレッドシート）")
        export_base_name = st.text_input(
            "出力名（例: 4月11日座席表）",
            value=default_daily_sheet_name(),
            key="layout_export_base_name",
            help=(
                "この名前を基準に、通常版は「（受講生視点）」、"
                "反転版は「（講師視点）」を自動で付与します。"
            ),
        )
        layout_payload_excel = build_layout_payload(
            current_rows,
            export_base_name,
            include_skill=False,
            table_layout_rows=active_layout_rows,
        )
        layout_payload_sheet = build_layout_payload(
            current_rows,
            export_base_name,
            include_skill=False,
            table_layout_rows=active_layout_rows,
        )
        xlsx_bytes = export_layouts_to_excel_bytes(
            normal_grid=layout_payload_excel["normal_grid"],
            mirrored_grid=layout_payload_excel["mirrored_grid"],
            normal_sheet_name=str(layout_payload_excel["normal_name"]),
            mirrored_sheet_name=str(layout_payload_excel["mirrored_name"]),
        )
        st.download_button(
            label="現在の結果をExcel(.xlsx)で出力（通常版+反転版）",
            data=xlsx_bytes,
            file_name=safe_download_filename(
                str(layout_payload_excel["normal_name"]), suffix=".xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_current_result_xlsx",
        )
        st.caption(
            "Excel/Googleスプレッドシート出力は氏名ふりがな・氏名・会社のみを記載し、"
            "スキルは出力しません。"
        )

        st.markdown("##### Googleスプレッドシートへ自動作成")
        spreadsheet_ref = st.text_input(
            "出力先スプレッドシートURL/ID",
            value=(
                "https://docs.google.com/spreadsheets/d/"
                "1XsHLia82yBlWYBC8ZfLJ6LARnKHI8OiGk_o78p5d8SI/edit"
            ),
            key="google_export_spreadsheet_ref",
        )
        template_sheet_name = st.text_input(
            "テンプレートシート名",
            value="座席表（テンプレ）",
            key="google_export_template_sheet_name",
        )
        service_account_file = st.file_uploader(
            "GoogleサービスアカウントJSON",
            type=["json"],
            key="google_service_account_json",
        )
        st.caption(
            "サービスアカウントの `client_email` を対象スプレッドシートに編集者として共有してから実行してください。"
        )
        if st.button("Googleスプレッドシートへ2種類を作成", key="export_to_google_sheet"):
            if service_account_file is None:
                st.error("GoogleサービスアカウントJSONを指定してください。")
            else:
                try:
                    publish_result = publish_layouts_to_google_sheets(
                        service_account_json_bytes=service_account_file.getvalue(),
                        spreadsheet_ref=spreadsheet_ref,
                        template_sheet_name=template_sheet_name,
                        normal_sheet_name=str(layout_payload_sheet["normal_name"]),
                        normal_grid=layout_payload_sheet["normal_grid"],
                        mirrored_sheet_name=str(layout_payload_sheet["mirrored_name"]),
                        mirrored_grid=layout_payload_sheet["mirrored_grid"],
                    )
                except Exception as exc:
                    st.error(f"スプレッドシート出力に失敗しました: {exc}")
                else:
                    st.success("スプレッドシートへ受講生視点・講師視点を作成しました。")
                    requested_template = str(publish_result.get("requested_template_sheet_name", ""))
                    used_template = str(publish_result.get("used_template_sheet_name", ""))
                    if requested_template and used_template and requested_template != used_template:
                        st.info(
                            "指定テンプレートが見つからなかったため、"
                            f"「{used_template}」をテンプレートとして使用しました。"
                        )
                    st.markdown(
                        f"- 通常版: [{publish_result['normal_sheet_name']}]"
                        f"({publish_result['normal_sheet_url']})"
                    )
                    st.markdown(
                        f"- 反転版: [{publish_result['mirrored_sheet_name']}]"
                        f"({publish_result['mirrored_sheet_url']})"
                    )

        st.caption("印刷はブラウザ標準機能（Ctrl+P / Cmd+P）を利用してください。")
        if st.button("この結果を履歴に保存", type="primary", key="save_current_history"):
            settings = st.session_state.get("current_settings", {})
            history_id = create_seating_history(
                target_week=str(current_target_week),
                settings=settings,
                total_score=float(st.session_state.get("current_score", 0.0)),
                overlap_rate=float(
                    st.session_state.get("current_metrics", {}).get("overlap_pair_rate", 0.0)
                ),
            )
            save_assignments(history_id, current_rows)
            st.success(f"履歴保存しました。履歴ID: {history_id}")


with tab_history:
    st.subheader("履歴")
    histories = list_histories(limit=300)
    if not histories:
        st.info("履歴はまだありません。")
    else:
        df_histories = pd.DataFrame(histories)
        st.dataframe(
            df_histories[
                ["id", "target_week", "assigned_students", "total_score", "overlap_rate", "created_at"]
            ],
            use_container_width=True,
            hide_index=True,
        )

        selected_history_id = st.selectbox(
            "詳細表示する履歴",
            options=[int(h["id"]) for h in histories],
            format_func=lambda hid: next(
                (
                    f"ID={h['id']} / {h['target_week']} / {h['assigned_students']}名"
                    for h in histories
                    if int(h["id"]) == int(hid)
                ),
                str(hid),
            ),
            key="selected_history_id",
        )

        history = get_history(int(selected_history_id))
        history_rows = get_history_rows(int(selected_history_id))
        if history and history_rows:
            st.write(
                f"対象週: **{history['target_week']}** / "
                f"スコア: **{float(history.get('total_score') or 0.0):.1f}** / "
                f"重複率: **{100 * float(history.get('overlap_rate') or 0.0):.1f}%**"
            )

            settings_json = history.get("settings_json")
            settings: dict[str, Any] = {}
            if settings_json:
                try:
                    settings = json.loads(settings_json)
                except json.JSONDecodeError:
                    settings = {}
            config = _config_from_settings(settings if settings else None)
            history_layout_rows = _coerce_layout_rows(
                settings.get("table_layout_rows") if isinstance(settings, dict) else None,
                config.table_count,
            )

            table_map = build_table_view(history_rows, config.table_count)
            for row_label, row_table_nos in zip(_LAYOUT_ROW_LABELS, history_layout_rows):
                st.markdown(f"##### {row_label}")
                if not row_table_nos:
                    st.caption("（設定なし）")
                    continue
                row_cols = st.columns(len(row_table_nos))
                for idx, table_no in enumerate(row_table_nos):
                    members = table_map.get(table_no, [])
                    with row_cols[idx]:
                        with st.container(border=True):
                            st.markdown(f"##### Table {table_no} ({len(members)}名)")
                            for member in members:
                                kana = str(member.get("name_kana", "")).strip()
                                prefix = f"{kana} / " if kana else ""
                                st.write(
                                    (
                                        f"- {prefix}{member['name']} / "
                                        f"{member['company']} / {skill_to_display(str(member['skill_level']))}"
                                    )
                                )

            history_displayed = {no for row in history_layout_rows for no in row}
            history_hidden = [
                table_no
                for table_no in range(1, config.table_count + 1)
                if table_no not in history_displayed
            ]
            if history_hidden:
                st.caption(
                    "この履歴で表示対象外のテーブル: "
                    + ", ".join(str(no) for no in history_hidden)
                )

            history_csv = rows_to_assignment_csv(
                rows=history_rows,
                target_week=str(history["target_week"]),
                history_id=int(selected_history_id),
            )
            st.download_button(
                label="この履歴をCSV出力",
                data=history_csv,
                file_name=f"seating_history_{selected_history_id}.csv",
                mime="text/csv",
                key="download_selected_history_csv",
            )

            if st.button("この履歴を結果タブに読み込む", key="load_history_to_result"):
                latest_history_id, previous_pairs, previous_table_map = get_previous_context()
                score, metrics = evaluate_rows(
                    rows=history_rows,
                    previous_pairs=previous_pairs,
                    previous_table_map=previous_table_map,
                    config=config,
                )
                _store_result(
                    rows=history_rows,
                    score=score,
                    metrics=metrics,
                    config=config,
                    target_week=str(history["target_week"]),
                    table_layout_rows=history_layout_rows,
                    previous_history_id=latest_history_id,
                    previous_pairs=previous_pairs,
                    previous_table_map=previous_table_map,
                )
                st.success("結果タブへ読み込みました。")
