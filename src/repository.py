from __future__ import annotations

import json
import unicodedata
from collections import defaultdict
from typing import Any

from .db import get_connection


def _reset_autoincrement(conn: Any, table_names: list[str]) -> None:
    seq_exists = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name='sqlite_sequence'"
    ).fetchone()
    if not seq_exists:
        return
    conn.executemany(
        "DELETE FROM sqlite_sequence WHERE name = ?",
        [(table_name,) for table_name in table_names],
    )


def _normalize_student_name(value: str) -> str:
    return " ".join(value.strip().replace("　", " ").split())


def _normalize_student_identity_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value or "")
    compact = normalized.strip().replace("　", " ").replace(" ", "")
    return compact.lower()


def _student_identity_key(name: str, company: str) -> tuple[str, str]:
    return (
        _normalize_student_identity_text(name),
        _normalize_student_identity_text(company),
    )


def list_students(search: str = "") -> list[dict[str, Any]]:
    query = """
        SELECT id, name, name_kana, company, skill_level, created_at, updated_at
        FROM students
    """
    params: tuple[Any, ...] = ()
    if search.strip():
        query += " WHERE name LIKE ? OR name_kana LIKE ? OR company LIKE ? "
        like = f"%{search.strip()}%"
        params = (like, like, like)
    query += " ORDER BY id ASC"

    with get_connection() as conn:
        rows = conn.execute(query, params).fetchall()
    return [dict(row) for row in rows]


def get_student(student_id: int) -> dict[str, Any] | None:
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT id, name, name_kana, company, skill_level, created_at, updated_at
            FROM students
            WHERE id = ?
            """,
            (student_id,),
        ).fetchone()
    return dict(row) if row else None


def create_student(name: str, company: str, skill_level: str, name_kana: str = "") -> int:
    with get_connection() as conn:
        cursor = conn.execute(
            """
            INSERT INTO students (name, name_kana, company, skill_level, created_at, updated_at)
            VALUES (?, ?, ?, ?, datetime('now'), datetime('now'))
            """,
            (name.strip(), name_kana.strip(), company.strip(), skill_level),
        )
        return int(cursor.lastrowid)


def update_student(
    student_id: int, name: str, company: str, skill_level: str, name_kana: str = ""
) -> None:
    with get_connection() as conn:
        conn.execute(
            """
            UPDATE students
            SET name = ?, name_kana = ?, company = ?, skill_level = ?, updated_at = datetime('now')
            WHERE id = ?
            """,
            (name.strip(), name_kana.strip(), company.strip(), skill_level, student_id),
        )


def delete_student(student_id: int) -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM students WHERE id = ?", (student_id,))


def delete_all_students() -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM students")
        _reset_autoincrement(conn, ["students"])


def delete_all_histories() -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM seating_histories")
        # seating_histories 削除時に assignments も消えるため両方リセット
        _reset_autoincrement(conn, ["seating_histories", "seating_assignments"])


def bulk_insert_students(rows: list[dict[str, str]]) -> int:
    if not rows:
        return 0

    with get_connection() as conn:
        existing_rows = conn.execute(
            "SELECT id, name, company, name_kana FROM students"
        ).fetchall()
        existing_by_key: dict[tuple[str, str], Any] = {}
        for row in existing_rows:
            key = _student_identity_key(str(row["name"]), str(row["company"]))
            if key not in existing_by_key:
                existing_by_key[key] = row
        existing_keys = set(existing_by_key.keys())
        seen_keys = set(existing_keys)

        values: list[tuple[str, str, str, str]] = []
        kana_updates: list[tuple[str, int]] = []
        for row in rows:
            name = row["name"].strip()
            name_kana = row.get("name_kana", "").strip()
            company = row["company"].strip()
            key = _student_identity_key(name, company)
            if key in existing_by_key:
                existing = existing_by_key[key]
                existing_kana = str(existing["name_kana"] or "").strip()
                # 既存データにふりがなが無い場合は、再インポートで補完する
                if name_kana and name_kana != existing_kana:
                    kana_updates.append((name_kana, int(existing["id"])))
                continue
            if key in seen_keys:
                continue
            seen_keys.add(key)
            values.append((name, name_kana, company, row["skill_level"]))

        if kana_updates:
            conn.executemany(
                """
                UPDATE students
                SET name_kana = ?, updated_at = datetime('now')
                WHERE id = ?
                """,
                kana_updates,
            )

        if not values:
            return 0

        conn.executemany(
            """
            INSERT INTO students (name, name_kana, company, skill_level, created_at, updated_at)
            VALUES (?, ?, ?, ?, datetime('now'), datetime('now'))
            """,
            values,
        )
    return len(values)


def bulk_update_student_skills_by_name(
    name_to_skill: dict[str, str],
) -> tuple[int, list[str], dict[str, int]]:
    if not name_to_skill:
        return 0, [], {}

    with get_connection() as conn:
        existing_rows = conn.execute("SELECT id, name FROM students").fetchall()
        index: dict[str, list[int]] = defaultdict(list)
        for row in existing_rows:
            normalized = _normalize_student_name(str(row["name"]))
            index[normalized].append(int(row["id"]))

        update_values: list[tuple[str, int]] = []
        unmatched_names: list[str] = []
        matched_name_counts: dict[str, int] = {}

        for raw_name, skill_level in name_to_skill.items():
            normalized_name = _normalize_student_name(raw_name)
            matched_ids = index.get(normalized_name, [])
            if not matched_ids:
                unmatched_names.append(raw_name)
                continue
            matched_name_counts[raw_name] = len(matched_ids)
            for student_id in matched_ids:
                update_values.append((skill_level, student_id))

        if update_values:
            conn.executemany(
                """
                UPDATE students
                SET skill_level = ?, updated_at = datetime('now')
                WHERE id = ?
                """,
                update_values,
            )

    return len(update_values), unmatched_names, matched_name_counts


def get_skill_distribution() -> dict[str, int]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT skill_level, COUNT(*) AS count
            FROM students
            GROUP BY skill_level
            """
        ).fetchall()
    distribution = {row["skill_level"]: int(row["count"]) for row in rows}
    return distribution


def create_seating_history(
    target_week: str,
    settings: dict[str, Any],
    total_score: float,
    overlap_rate: float,
) -> int:
    with get_connection() as conn:
        cursor = conn.execute(
            """
            INSERT INTO seating_histories (target_week, settings_json, total_score, overlap_rate, created_at)
            VALUES (?, ?, ?, ?, datetime('now'))
            """,
            (target_week, json.dumps(settings, ensure_ascii=False), total_score, overlap_rate),
        )
        return int(cursor.lastrowid)


def save_assignments(history_id: int, rows: list[dict[str, Any]]) -> None:
    values = [(history_id, int(row["table_no"]), int(row["student_id"])) for row in rows]
    with get_connection() as conn:
        conn.executemany(
            """
            INSERT INTO seating_assignments (seating_history_id, table_no, student_id)
            VALUES (?, ?, ?)
            """,
            values,
        )


def list_histories(limit: int = 200) -> list[dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT
                h.id,
                h.target_week,
                h.total_score,
                h.overlap_rate,
                h.created_at,
                COUNT(a.id) AS assigned_students
            FROM seating_histories h
            LEFT JOIN seating_assignments a ON a.seating_history_id = h.id
            GROUP BY h.id
            ORDER BY h.id DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
    return [dict(row) for row in rows]


def get_history_rows(history_id: int) -> list[dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT
                a.table_no,
                s.id AS student_id,
                s.name,
                s.name_kana,
                s.company,
                s.skill_level
            FROM seating_assignments a
            INNER JOIN students s ON s.id = a.student_id
            WHERE a.seating_history_id = ?
            ORDER BY a.table_no ASC, s.id ASC
            """,
            (history_id,),
        ).fetchall()
    return [dict(row) for row in rows]


def get_history(history_id: int) -> dict[str, Any] | None:
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT id, target_week, settings_json, total_score, overlap_rate, created_at
            FROM seating_histories
            WHERE id = ?
            """,
            (history_id,),
        ).fetchone()
    return dict(row) if row else None


def get_latest_history_id() -> int | None:
    with get_connection() as conn:
        row = conn.execute("SELECT id FROM seating_histories ORDER BY id DESC LIMIT 1").fetchone()
    return int(row["id"]) if row else None


def get_previous_context() -> tuple[int | None, set[tuple[int, int]], dict[int, int]]:
    """
    Returns:
        latest_history_id,
        previous pair set (sorted student-id tuple),
        previous student_id -> table_no map
    """
    history_id = get_latest_history_id()
    if history_id is None:
        return None, set(), {}

    rows = get_history_rows(history_id)
    table_members: dict[int, list[int]] = defaultdict(list)
    student_to_table: dict[int, int] = {}
    for row in rows:
        sid = int(row["student_id"])
        table_no = int(row["table_no"])
        table_members[table_no].append(sid)
        student_to_table[sid] = table_no

    pair_set: set[tuple[int, int]] = set()
    for members in table_members.values():
        for i in range(len(members)):
            for j in range(i + 1, len(members)):
                a = members[i]
                b = members[j]
                if a < b:
                    pair_set.add((a, b))
                else:
                    pair_set.add((b, a))

    return history_id, pair_set, student_to_table
