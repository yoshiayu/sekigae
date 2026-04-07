from __future__ import annotations

SKILL_LEVELS = ("高い", "並", "低い", "ヤバい")

SKILL_COLORS = {
    "高い": "#2B6CB0",
    "並": "#2F855A",
    "低い": "#B7791F",
    "ヤバい": "#C53030",
}

CSV_COLUMNS = ("name", "company", "skill_level")

DEFAULT_TABLE_COUNT = 13
DEFAULT_MAX_PER_TABLE = 6
DEFAULT_ATTEMPTS = 300

DEFAULT_WEIGHTS = {
    "company": 26.0,
    "previous": 14.0,
    "skill": 24.0,
    "size": 2.0,
    "randomness": 1.0,
}

