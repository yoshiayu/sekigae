from __future__ import annotations

import math
import random
from collections import Counter, defaultdict
from dataclasses import dataclass
from typing import Any

from .constants import (
    DEFAULT_ATTEMPTS,
    DEFAULT_MAX_PER_TABLE,
    DEFAULT_TABLE_COUNT,
    DEFAULT_WEIGHTS,
    SKILL_LEVELS,
)


@dataclass(frozen=True)
class Student:
    id: int
    name: str
    company: str
    skill_level: str


@dataclass(frozen=True)
class SeatingConfig:
    table_count: int = DEFAULT_TABLE_COUNT
    max_per_table: int = DEFAULT_MAX_PER_TABLE
    attempts: int = DEFAULT_ATTEMPTS
    company_weight: float = DEFAULT_WEIGHTS["company"]
    previous_weight: float = DEFAULT_WEIGHTS["previous"]
    skill_weight: float = DEFAULT_WEIGHTS["skill"]
    size_weight: float = DEFAULT_WEIGHTS["size"]
    randomness_weight: float = DEFAULT_WEIGHTS["randomness"]
    seed: int | None = None


def _pair_key(a: int, b: int) -> tuple[int, int]:
    return (a, b) if a < b else (b, a)


def _expected_table_sizes(total_students: int, table_count: int) -> list[int]:
    base = total_students // table_count
    rem = total_students % table_count
    return [base + 1 if i < rem else base for i in range(table_count)]


def _students_to_rows(tables: list[list[Student]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for table_idx, members in enumerate(tables, start=1):
        for student in members:
            rows.append(
                {
                    "table_no": table_idx,
                    "student_id": student.id,
                    "name": student.name,
                    "company": student.company,
                    "skill_level": student.skill_level,
                }
            )
    rows.sort(key=lambda r: (int(r["table_no"]), int(r["student_id"])))
    return rows


def _evaluate_assignment(
    rows: list[dict[str, Any]],
    previous_pairs: set[tuple[int, int]],
    previous_table_map: dict[int, int],
    config: SeatingConfig,
) -> tuple[float, dict[str, Any]]:
    table_members: dict[int, list[dict[str, Any]]] = {i: [] for i in range(1, config.table_count + 1)}
    for row in rows:
        table_members[int(row["table_no"])].append(row)

    total_students = len(rows)
    expected_sizes = _expected_table_sizes(total_students, config.table_count)
    avg_skill_targets = {
        skill: sum(1 for row in rows if row["skill_level"] == skill) / float(config.table_count)
        for skill in SKILL_LEVELS
    }

    company_collision_total = 0
    previous_pair_hits = 0
    skill_deviation_total = 0.0
    size_deviation_total = 0.0
    current_pairs: set[tuple[int, int]] = set()

    table_metrics: list[dict[str, Any]] = []

    for table_no in range(1, config.table_count + 1):
        members = table_members[table_no]
        company_counter = Counter(m["company"] for m in members)
        skill_counter = Counter(m["skill_level"] for m in members)
        duplicate_companies = {
            company: count for company, count in company_counter.items() if count > 1
        }
        company_collisions = sum(count - 1 for count in company_counter.values() if count > 1)
        company_collision_total += company_collisions

        for i in range(len(members)):
            sid_i = int(members[i]["student_id"])
            for j in range(i + 1, len(members)):
                sid_j = int(members[j]["student_id"])
                pair = _pair_key(sid_i, sid_j)
                current_pairs.add(pair)
                if pair in previous_pairs:
                    previous_pair_hits += 1

        for skill in SKILL_LEVELS:
            skill_deviation_total += abs(skill_counter.get(skill, 0) - avg_skill_targets[skill])

        target_size = expected_sizes[table_no - 1]
        size_deviation_total += abs(len(members) - target_size)

        table_metrics.append(
            {
                "table_no": table_no,
                "size": len(members),
                "skill_counts": {skill: skill_counter.get(skill, 0) for skill in SKILL_LEVELS},
                "duplicate_companies": duplicate_companies,
                "previous_pair_hits": sum(
                    1
                    for i in range(len(members))
                    for j in range(i + 1, len(members))
                    if _pair_key(int(members[i]["student_id"]), int(members[j]["student_id"]))
                    in previous_pairs
                ),
            }
        )

    same_table_student_count = 0
    if previous_table_map:
        for row in rows:
            sid = int(row["student_id"])
            if previous_table_map.get(sid) == int(row["table_no"]):
                same_table_student_count += 1
    same_table_rate = (
        same_table_student_count / float(total_students) if total_students > 0 and previous_table_map else 0.0
    )

    overlap_pair_rate = (
        (len(current_pairs & previous_pairs) / float(len(current_pairs)))
        if current_pairs and previous_pairs
        else 0.0
    )

    total_score = (
        company_collision_total * config.company_weight
        + previous_pair_hits * config.previous_weight
        + skill_deviation_total * config.skill_weight
        + size_deviation_total * config.size_weight
    )

    metrics = {
        "company_collision_total": company_collision_total,
        "previous_pair_hits": previous_pair_hits,
        "skill_deviation_total": round(skill_deviation_total, 3),
        "size_deviation_total": round(size_deviation_total, 3),
        "overlap_pair_rate": overlap_pair_rate,
        "same_table_student_rate": same_table_rate,
        "table_metrics": table_metrics,
        "total_students": total_students,
    }
    return float(total_score), metrics


def _single_attempt(
    students: list[Student],
    previous_pairs: set[tuple[int, int]],
    config: SeatingConfig,
    rng: random.Random,
) -> tuple[list[dict[str, Any]], float]:
    table_count = config.table_count
    max_per_table = config.max_per_table
    total_students = len(students)

    tables: list[list[Student]] = [[] for _ in range(table_count)]
    table_company_counts: list[defaultdict[str, int]] = [defaultdict(int) for _ in range(table_count)]
    table_skill_counts: list[defaultdict[str, int]] = [defaultdict(int) for _ in range(table_count)]

    expected_sizes = _expected_table_sizes(total_students, table_count)
    rng.shuffle(expected_sizes)

    skill_counts = Counter(s.skill_level for s in students)
    skill_targets = {skill: skill_counts.get(skill, 0) / float(table_count) for skill in SKILL_LEVELS}
    company_counts = Counter(s.company for s in students)

    order = students[:]
    rng.shuffle(order)
    order.sort(
        key=lambda s: (
            skill_counts.get(s.skill_level, 0),  # レアなスキルを先に置く
            -company_counts.get(s.company, 0),  # 人数が多い会社を先に分散
            rng.random(),
        )
    )

    for student in order:
        candidates: list[tuple[float, int]] = []
        for table_idx in range(table_count):
            members = tables[table_idx]
            if len(members) >= max_per_table:
                continue

            company_penalty = table_company_counts[table_idx][student.company] * config.company_weight

            prev_hits = 0
            for other in members:
                if _pair_key(student.id, other.id) in previous_pairs:
                    prev_hits += 1
            previous_penalty = prev_hits * config.previous_weight

            target_skill = skill_targets[student.skill_level]
            before_skill = abs(table_skill_counts[table_idx][student.skill_level] - target_skill)
            after_skill = abs(table_skill_counts[table_idx][student.skill_level] + 1 - target_skill)
            skill_penalty = (after_skill - before_skill) * config.skill_weight

            target_size = expected_sizes[table_idx]
            before_size = abs(len(members) - target_size)
            after_size = abs(len(members) + 1 - target_size)
            size_penalty = (after_size - before_size) * config.size_weight

            jitter = rng.random() * config.randomness_weight
            score = company_penalty + previous_penalty + skill_penalty + size_penalty + jitter
            candidates.append((score, table_idx))

        if not candidates:
            return [], math.inf

        candidates.sort(key=lambda x: x[0])
        top_k = candidates[: min(3, len(candidates))]
        ranked_weights = [1.0 / (i + 1) for i in range(len(top_k))]
        _, chosen_table = rng.choices(top_k, weights=ranked_weights, k=1)[0]

        tables[chosen_table].append(student)
        table_company_counts[chosen_table][student.company] += 1
        table_skill_counts[chosen_table][student.skill_level] += 1

    rows = _students_to_rows(tables)
    score, _ = _evaluate_assignment(rows, previous_pairs, previous_table_map={}, config=config)
    return rows, score


def generate_best_assignment(
    students: list[Student],
    previous_pairs: set[tuple[int, int]],
    previous_table_map: dict[int, int],
    config: SeatingConfig,
) -> dict[str, Any]:
    capacity = config.table_count * config.max_per_table
    if len(students) > capacity:
        raise ValueError(
            f"受講生数が上限を超えています。students={len(students)}, capacity={capacity}"
        )

    if not students:
        return {
            "rows": [],
            "score": 0.0,
            "metrics": {
                "company_collision_total": 0,
                "previous_pair_hits": 0,
                "skill_deviation_total": 0.0,
                "size_deviation_total": 0.0,
                "overlap_pair_rate": 0.0,
                "same_table_student_rate": 0.0,
                "table_metrics": [],
                "total_students": 0,
            },
        }

    base_rng = random.Random(config.seed)
    best_rows: list[dict[str, Any]] = []
    best_score = math.inf

    for _ in range(max(1, config.attempts)):
        attempt_rng = random.Random(base_rng.randint(1, 10**9))
        rows, rough_score = _single_attempt(students, previous_pairs, config, attempt_rng)
        if not rows:
            continue
        if rough_score < best_score:
            best_rows = rows
            best_score = rough_score

    if not best_rows:
        raise RuntimeError("席替え計算に失敗しました。設定を見直してください。")

    final_score, metrics = _evaluate_assignment(best_rows, previous_pairs, previous_table_map, config)
    return {"rows": best_rows, "score": final_score, "metrics": metrics}


def validate_manual_rows(
    rows: list[dict[str, Any]], table_count: int, max_per_table: int
) -> list[str]:
    errors: list[str] = []
    if any(int(row["table_no"]) < 1 or int(row["table_no"]) > table_count for row in rows):
        errors.append(f"table_no は 1〜{table_count} の範囲で入力してください。")

    table_sizes = Counter(int(row["table_no"]) for row in rows)
    over_capacity = [table_no for table_no, size in table_sizes.items() if size > max_per_table]
    if over_capacity:
        over_capacity_str = ", ".join(str(no) for no in sorted(over_capacity))
        errors.append(
            f"テーブル上限({max_per_table}名)を超えているテーブルがあります: {over_capacity_str}"
        )
    return errors


def build_table_view(rows: list[dict[str, Any]], table_count: int) -> dict[int, list[dict[str, Any]]]:
    grouped: dict[int, list[dict[str, Any]]] = {table_no: [] for table_no in range(1, table_count + 1)}
    for row in rows:
        grouped[int(row["table_no"])].append(row)
    for members in grouped.values():
        members.sort(key=lambda r: int(r["student_id"]))
    return grouped


def evaluate_rows(
    rows: list[dict[str, Any]],
    previous_pairs: set[tuple[int, int]],
    previous_table_map: dict[int, int],
    config: SeatingConfig,
) -> tuple[float, dict[str, Any]]:
    return _evaluate_assignment(rows, previous_pairs, previous_table_map, config)
