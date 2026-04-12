from __future__ import annotations

SKILL_LEVELS = ("高い", "並", "低い", "ヤバい")
SKILL_DISPLAY_LABELS = {
    "高い": "高",
    "並": "中",
    "低い": "やや低",
    "ヤバい": "低",
}
SKILL_DISPLAY_ORDER = tuple(SKILL_DISPLAY_LABELS[level] for level in SKILL_LEVELS)
DISPLAY_TO_INTERNAL_SKILL = {label: level for level, label in SKILL_DISPLAY_LABELS.items()}

SKILL_COLORS = {
    "高い": "#2B6CB0",
    "並": "#2F855A",
    "低い": "#B7791F",
    "ヤバい": "#C53030",
}

CSV_COLUMNS = ("name", "company", "skill_level")

DEFAULT_TABLE_COUNT = 13
DEFAULT_MIN_PER_TABLE = 5
DEFAULT_MAX_PER_TABLE = 6
DEFAULT_ATTEMPTS = 300

DEFAULT_WEIGHTS = {
    "company": 26.0,
    "previous": 14.0,
    "skill": 32.0,
    "size": 2.0,
    "randomness": 1.0,
}


def skill_to_display(skill_level: str) -> str:
    return SKILL_DISPLAY_LABELS.get(skill_level, skill_level)


def skill_to_internal(skill_level: str) -> str:
    return DISPLAY_TO_INTERNAL_SKILL.get(skill_level, skill_level)
