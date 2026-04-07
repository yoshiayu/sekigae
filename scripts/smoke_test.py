from __future__ import annotations

from collections import Counter
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.constants import DEFAULT_MAX_PER_TABLE, DEFAULT_TABLE_COUNT
from src.seating import SeatingConfig, Student, generate_best_assignment


def make_dummy_students(total: int = 73) -> list[Student]:
    skills = ["高い"] * 15 + ["並"] * 36 + ["低い"] * 15 + ["ヤバい"] * 7
    companies = [f"会社{chr(65 + (i % 13))}" for i in range(total)]
    students: list[Student] = []
    for i in range(total):
        students.append(
            Student(
                id=i + 1,
                name=f"受講生{i + 1:03d}",
                company=companies[i],
                skill_level=skills[i],
            )
        )
    return students


def main() -> None:
    students = make_dummy_students(73)
    config = SeatingConfig(
        table_count=DEFAULT_TABLE_COUNT,
        max_per_table=DEFAULT_MAX_PER_TABLE,
        attempts=200,
    )
    result = generate_best_assignment(
        students=students,
        previous_pairs=set(),
        previous_table_map={},
        config=config,
    )

    rows = result["rows"]
    assert len(rows) == 73, "全員が配置されていません"
    table_sizes = Counter(row["table_no"] for row in rows)
    assert all(size <= DEFAULT_MAX_PER_TABLE for size in table_sizes.values()), "6名超過あり"
    print("smoke test passed")
    print(f"score={result['score']:.2f}")


if __name__ == "__main__":
    main()
