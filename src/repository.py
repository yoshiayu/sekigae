from __future__ import annotations

import json
from collections import defaultdict
from typing import Any

from .db import get_connection


def list_students(search: str = "") -> list[dict[str, Any]]:
    query = """
        SELECT id, name, company, skill_level, created_at, updated_at
        FROM students
    """
    params: tuple[Any, ...] = ()
    if search.strip():
        query += " WHERE name LIKE ? OR company LIKE ? "
        like = f"%{search.strip()}%"
        params = (like, like)
    query += " ORDER BY id ASC"

    with get_connection() as conn:
        rows = conn.execute(query, params).fetchall()
    return [dict(row) for row in rows]


def get_student(student_id: int) -> dict[str, Any] | None:
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT id, name, company, skill_level, created_at, updated_at
            FROM students
            WHERE id = ?
            """,
            (student_id,),
        ).fetchone()
    return dict(row) if row else None


def create_student(name: str, company: str, skill_level: str) -> int:
    with get_connection() as conn:
        cursor = conn.execute(
            """
            INSERT INTO students (name, company, skill_level, created_at, updated_at)
            VALUES (?, ?, ?, datetime('now'), datetime('now'))
            """,
            (name.strip(), company.strip(), skill_level),
        )
        return int(cursor.lastrowid)


def update_student(student_id: int, name: str, company: str, skill_level: str) -> None:
    with get_connection() as conn:
        conn.execute(
            """
            UPDATE students
            SET name = ?, company = ?, skill_level = ?, updated_at = datetime('now')
            WHERE id = ?
            """,
            (name.strip(), company.strip(), skill_level, student_id),
        )


def delete_student(student_id: int) -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM students WHERE id = ?", (student_id,))


def delete_all_students() -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM students")


def bulk_insert_students(rows: list[dict[str, str]]) -> int:
    if not rows:
        return 0
    values = [(row["name"].strip(), row["company"].strip(), row["skill_level"]) for row in rows]
    with get_connection() as conn:
        conn.executemany(
            """
            INSERT INTO students (name, company, skill_level, created_at, updated_at)
            VALUES (?, ?, ?, datetime('now'), datetime('now'))
            """,
            values,
        )
    return len(values)


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

