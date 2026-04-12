from __future__ import annotations

import sqlite3
from pathlib import Path

DB_PATH = Path(__file__).resolve().parents[1] / "seating_app.db"


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def init_db() -> None:
    with get_connection() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS students (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                name_kana TEXT NOT NULL DEFAULT '',
                company TEXT NOT NULL,
                skill_level TEXT NOT NULL CHECK (skill_level IN ('高い', '並', '低い', 'ヤバい')),
                created_at TEXT NOT NULL DEFAULT (datetime('now')),
                updated_at TEXT NOT NULL DEFAULT (datetime('now'))
            );

            CREATE TABLE IF NOT EXISTS seating_histories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                target_week TEXT NOT NULL,
                settings_json TEXT,
                total_score REAL,
                overlap_rate REAL,
                created_at TEXT NOT NULL DEFAULT (datetime('now'))
            );

            CREATE TABLE IF NOT EXISTS seating_assignments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                seating_history_id INTEGER NOT NULL,
                table_no INTEGER NOT NULL,
                student_id INTEGER NOT NULL,
                FOREIGN KEY (seating_history_id) REFERENCES seating_histories(id) ON DELETE CASCADE,
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE
            );

            CREATE UNIQUE INDEX IF NOT EXISTS idx_assignment_unique_history_student
                ON seating_assignments (seating_history_id, student_id);
            CREATE INDEX IF NOT EXISTS idx_assignment_history_table
                ON seating_assignments (seating_history_id, table_no);
            CREATE INDEX IF NOT EXISTS idx_student_company ON students (company);
            CREATE INDEX IF NOT EXISTS idx_student_skill ON students (skill_level);
            """
        )

        # Lightweight migration for existing DBs.
        student_columns = {
            str(row["name"])
            for row in conn.execute("PRAGMA table_info(students)").fetchall()
        }
        if "name_kana" not in student_columns:
            conn.execute("ALTER TABLE students ADD COLUMN name_kana TEXT NOT NULL DEFAULT ''")
