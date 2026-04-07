from __future__ import annotations

import json
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from src.constants import (
    DEFAULT_ATTEMPTS,
    DEFAULT_MAX_PER_TABLE,
    DEFAULT_TABLE_COUNT,
    DEFAULT_WEIGHTS,
    SKILL_COLORS,
    SKILL_LEVELS,
)
from src.csv_service import parse_students_csv, rows_to_assignment_csv, template_csv_text
from src.db import init_db
from src.repository import (
    bulk_insert_students,
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
    evaluate_rows,
    generate_best_assignment,
    validate_manual_rows,
)


st.set_page_config(page_title="受講生席替えアプリ", page_icon="席", layout="wide")
init_db()


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
        chunks.append(f"{skill}: {count}名 ({ratio:.1f}%)")
    return " / ".join(chunks)


def _rows_to_students(rows: list[dict[str, Any]]) -> list[Student]:
    return [
        Student(
            id=int(row["id"]),
            name=str(row["name"]),
            company=str(row["company"]),
            skill_level=str(row["skill_level"]),
        )
        for row in rows
    ]


def _skill_badge(skill: str) -> str:
    color = SKILL_COLORS.get(skill, "#4A5568")
    return (
        f"<span style='background:{color};color:#fff;border-radius:999px;"
        f"padding:2px 8px;font-size:12px;'>{skill}</span>"
    )


def _make_config_from_state() -> SeatingConfig:
    settings = st.session_state.get("current_settings")
    if isinstance(settings, dict):
        return SeatingConfig(**settings)
    return SeatingConfig()


def _store_result(
    rows: list[dict[str, Any]],
    score: float,
    metrics: dict[str, Any],
    config: SeatingConfig,
    target_week: str,
    previous_history_id: int | None,
    previous_pairs: set[tuple[int, int]],
    previous_table_map: dict[int, int],
) -> None:
    st.session_state["current_rows"] = rows
    st.session_state["current_score"] = score
    st.session_state["current_metrics"] = metrics
    st.session_state["current_target_week"] = target_week
    st.session_state["current_settings"] = {
        "table_count": config.table_count,
        "max_per_table": config.max_per_table,
        "attempts": config.attempts,
        "company_weight": config.company_weight,
        "previous_weight": config.previous_weight,
        "skill_weight": config.skill_weight,
        "size_weight": config.size_weight,
        "randomness_weight": config.randomness_weight,
        "seed": config.seed,
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


st.title("受講生席替えアプリ（MVP）")
st.caption("毎週運用を想定した、スキル分散 + 会社分散 + 前回同席回避 の自動席替え")

tab_students, tab_run, tab_result, tab_history = st.tabs(
    ["受講生管理", "席替え実行", "席替え結果", "履歴"]
)


with tab_students:
    st.subheader("受講生一覧")
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
        st.dataframe(
            df_students[["id", "name", "company", "skill_level", "updated_at"]],
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
            add_company = st.text_input("会社名", key="add_company")
            add_skill = st.selectbox("スキル", options=list(SKILL_LEVELS), key="add_skill")
            add_submit = st.form_submit_button("追加")
        if add_submit:
            if not add_name.strip() or not add_company.strip():
                st.error("氏名・会社名は必須です。")
            else:
                create_student(add_name, add_company, add_skill)
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
                edit_company = st.text_input(
                    "会社名",
                    value=str(selected["company"]),
                    key=f"edit_company_{selected_id}",
                )
                edit_skill = st.selectbox(
                    "スキル",
                    options=list(SKILL_LEVELS),
                    index=list(SKILL_LEVELS).index(str(selected["skill_level"])),
                    key=f"edit_skill_{selected_id}",
                )
                edit_submit = st.form_submit_button("更新")

            if edit_submit:
                if not edit_name.strip() or not edit_company.strip():
                    st.error("氏名・会社名は必須です。")
                else:
                    update_student(int(selected_id), edit_name, edit_company, edit_skill)
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
    capacity = DEFAULT_TABLE_COUNT * DEFAULT_MAX_PER_TABLE

    col_info1, col_info2, col_info3 = st.columns(3)
    col_info1.metric("登録人数", f"{student_count}名")
    col_info2.metric("総席数", f"{capacity}席")
    col_info3.metric("空席", f"{capacity - student_count}席")
    st.write(_skill_distribution_text(distribution))

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
            "スキル分散の強さ", 1.0, 60.0, float(DEFAULT_WEIGHTS["skill"]), 1.0
        )
        size_weight = st.slider("人数均等化の強さ", 0.0, 10.0, float(DEFAULT_WEIGHTS["size"]), 0.5)
    with col_w3:
        randomness_weight = st.slider(
            "ランダム性", 0.0, 5.0, float(DEFAULT_WEIGHTS["randomness"]), 0.1
        )
        attempts = st.slider("試行回数", 100, 500, DEFAULT_ATTEMPTS, 50)
        seed_text = st.text_input("乱数シード（任意）", value="", key="seed_text")

    run_disabled = student_count == 0 or student_count > capacity
    if st.button("席替えを実行", type="primary", disabled=run_disabled, key="run_seating"):
        try:
            seed_value = int(seed_text) if seed_text.strip() else None
        except ValueError:
            st.error("乱数シードは整数で入力してください。")
            seed_value = None
        else:
            previous_history_id, previous_pairs, previous_table_map = get_previous_context()
            config = SeatingConfig(
                table_count=DEFAULT_TABLE_COUNT,
                max_per_table=DEFAULT_MAX_PER_TABLE,
                attempts=attempts,
                company_weight=company_weight,
                previous_weight=previous_weight,
                skill_weight=skill_weight,
                size_weight=size_weight,
                randomness_weight=randomness_weight,
                seed=seed_value,
            )
            result = generate_best_assignment(
                students=_rows_to_students(students_rows),
                previous_pairs=previous_pairs,
                previous_table_map=previous_table_map,
                config=config,
            )
            _store_result(
                rows=result["rows"],
                score=float(result["score"]),
                metrics=result["metrics"],
                config=config,
                target_week=target_week,
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

        st.markdown("#### 手動調整（table_no を直接変更）")
        df_current = pd.DataFrame(current_rows)
        edited_df = st.data_editor(
            df_current[["table_no", "student_id", "name", "company", "skill_level"]],
            hide_index=True,
            use_container_width=True,
            disabled=["student_id", "name", "company", "skill_level"],
            column_config={
                "table_no": st.column_config.NumberColumn(
                    "table_no", min_value=1, max_value=current_config.table_count, step=1
                )
            },
            key="result_editor",
        )

        if st.button("手動調整を反映", key="apply_manual_changes"):
            updated_rows = edited_df.to_dict(orient="records")
            errors = validate_manual_rows(
                rows=updated_rows,
                table_count=current_config.table_count,
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

        st.markdown("#### テーブル別表示")
        table_map = build_table_view(current_rows, current_config.table_count)
        table_metrics = {
            int(m["table_no"]): m for m in st.session_state["current_metrics"].get("table_metrics", [])
        }

        columns = st.columns(3)
        for table_no in range(1, current_config.table_count + 1):
            col = columns[(table_no - 1) % 3]
            members = table_map[table_no]
            metric = table_metrics.get(table_no, {})

            with col:
                st.markdown(f"##### Table {table_no} ({len(members)}名)")
                if not members:
                    st.caption("未配置")
                for member in members:
                    badge = _skill_badge(str(member["skill_level"]))
                    st.markdown(
                        (
                            f"<div style='padding:6px 8px;border:1px solid #E2E8F0;"
                            f"border-radius:8px;margin-bottom:6px;'>"
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
            config = SeatingConfig(**settings) if settings else SeatingConfig()

            table_map = build_table_view(history_rows, config.table_count)
            cols = st.columns(3)
            for table_no in range(1, config.table_count + 1):
                col = cols[(table_no - 1) % 3]
                members = table_map[table_no]
                with col:
                    st.markdown(f"##### Table {table_no} ({len(members)}名)")
                    for member in members:
                        st.write(f"- {member['name']} / {member['company']} / {member['skill_level']}")

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
                    previous_history_id=latest_history_id,
                    previous_pairs=previous_pairs,
                    previous_table_map=previous_table_map,
                )
                st.success("結果タブへ読み込みました。")
