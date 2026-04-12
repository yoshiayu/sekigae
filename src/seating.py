from __future__ import annotations

import math
import random
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from typing import Any

from .constants import (
    DEFAULT_ATTEMPTS,
    DEFAULT_MAX_PER_TABLE,
    DEFAULT_MIN_PER_TABLE,
    DEFAULT_TABLE_COUNT,
    DEFAULT_WEIGHTS,
    SKILL_LEVELS,
)

_SKILL_PRIORITY = {skill: idx for idx, skill in enumerate(SKILL_LEVELS)}
_SOFT_AVOID_PAIR = frozenset(("並", "ヤバい"))


@dataclass(frozen=True)
class Student:
    id: int
    name: str
    company: str
    skill_level: str
    name_kana: str = ""


@dataclass(frozen=True)
class SeatingConfig:
    table_count: int = DEFAULT_TABLE_COUNT
    min_per_table: int = DEFAULT_MIN_PER_TABLE
    max_per_table: int = DEFAULT_MAX_PER_TABLE
    attempts: int = DEFAULT_ATTEMPTS
    company_weight: float = DEFAULT_WEIGHTS["company"]
    previous_weight: float = DEFAULT_WEIGHTS["previous"]
    skill_weight: float = DEFAULT_WEIGHTS["skill"]
    size_weight: float = DEFAULT_WEIGHTS["size"]
    randomness_weight: float = DEFAULT_WEIGHTS["randomness"]
    seed: int | None = None


def _normalize_identity_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value or "")
    compact = normalized.strip().replace("　", " ").replace(" ", "")
    return compact.lower()


def _normalize_company_key(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value or "")
    compact = normalized.strip().replace("　", " ").replace(" ", "")
    compact = compact.replace("株式会社", "(株)").replace("有限会社", "(有)")
    return compact.lower()


def _student_identity_key(student: Student) -> tuple[str, str]:
    return (
        _normalize_identity_text(student.name),
        _normalize_identity_text(student.company),
    )


def _row_identity_key(row: dict[str, Any]) -> tuple[str, str]:
    return (
        _normalize_identity_text(str(row.get("name", ""))),
        _normalize_identity_text(str(row.get("company", ""))),
    )


def _prefer_duplicate_candidate(existing: Student, candidate: Student) -> bool:
    existing_has_kana = bool(existing.name_kana.strip())
    candidate_has_kana = bool(candidate.name_kana.strip())
    if candidate_has_kana != existing_has_kana:
        return candidate_has_kana
    return int(candidate.id) > int(existing.id)


def _dedupe_students_for_assignment(students: list[Student]) -> list[Student]:
    if not students:
        return []

    # Step 1: student_id の重複を除去（補足情報が多い行を優先）
    by_id: dict[int, Student] = {}
    id_order: list[int] = []
    for student in students:
        student_id = int(student.id)
        if student_id not in by_id:
            by_id[student_id] = student
            id_order.append(student_id)
            continue
        if _prefer_duplicate_candidate(by_id[student_id], student):
            by_id[student_id] = student
    unique_by_id = [by_id[student_id] for student_id in id_order]

    # Step 2: 同一受講生（氏名+会社）の重複を除去
    by_identity: dict[tuple[str, str], Student] = {}
    identity_order: list[tuple[str, str]] = []
    for student in unique_by_id:
        identity_key = _student_identity_key(student)
        if identity_key not in by_identity:
            by_identity[identity_key] = student
            identity_order.append(identity_key)
            continue
        if _prefer_duplicate_candidate(by_identity[identity_key], student):
            by_identity[identity_key] = student
    return [by_identity[key] for key in identity_order]


def _find_duplicate_student_ids(rows: list[dict[str, Any]]) -> list[int]:
    counter: Counter[int] = Counter()
    for row in rows:
        try:
            student_id = int(row.get("student_id", 0))
        except (TypeError, ValueError):
            student_id = 0
        if student_id > 0:
            counter[student_id] += 1
    return sorted(student_id for student_id, count in counter.items() if count > 1)


def _find_duplicate_row_identities(rows: list[dict[str, Any]]) -> list[str]:
    counter: Counter[tuple[str, str]] = Counter()
    display: dict[tuple[str, str], tuple[str, str]] = {}
    for row in rows:
        identity_key = _row_identity_key(row)
        counter[identity_key] += 1
        if identity_key not in display:
            display[identity_key] = (
                str(row.get("name", "")).strip(),
                str(row.get("company", "")).strip(),
            )

    duplicates: list[str] = []
    for identity_key, count in counter.items():
        if count <= 1:
            continue
        name, company = display.get(identity_key, ("", ""))
        duplicates.append(f"{name} ({company}) x{count}")
    duplicates.sort()
    return duplicates


def _validate_no_duplicate_assignment_rows(
    rows: list[dict[str, Any]], expected_total: int | None = None
) -> None:
    duplicate_ids = _find_duplicate_student_ids(rows)
    if duplicate_ids:
        preview = ", ".join(str(student_id) for student_id in duplicate_ids[:10])
        suffix = " ..." if len(duplicate_ids) > 10 else ""
        raise RuntimeError(
            "配席結果に同一受講生IDの重複があります。"
            f" student_id={preview}{suffix}"
        )

    duplicate_identities = _find_duplicate_row_identities(rows)
    if duplicate_identities:
        preview = " / ".join(duplicate_identities[:6])
        suffix = " ..." if len(duplicate_identities) > 6 else ""
        raise RuntimeError(
            "配席結果に同一受講生（氏名+会社）の重複があります。"
            f" {preview}{suffix}"
        )

    if expected_total is not None and len(rows) != expected_total:
        raise RuntimeError(
            "配席結果の人数が不整合です。"
            f" expected={expected_total}, actual={len(rows)}"
        )


def _pair_key(a: int, b: int) -> tuple[int, int]:
    return (a, b) if a < b else (b, a)


def _skill_group(skill_level: str) -> str:
    # Skill grouping now uses exact levels because seating constraints
    # depend on adjacent skill distance.
    return skill_level


def _is_forbidden_skill_pair(skill_a: str, skill_b: str) -> bool:
    # Hard constraints:
    # - 高×やや低 は禁止
    # - 高×低 は禁止
    pair = frozenset((skill_a, skill_b))
    return pair in {frozenset(("高い", "低い")), frozenset(("高い", "ヤバい"))}


def _is_soft_avoid_pair(skill_a: str, skill_b: str) -> bool:
    return frozenset((skill_a, skill_b)) == _SOFT_AVOID_PAIR


def _bounded_balanced_table_sizes(
    total_students: int, table_count: int, min_per_table: int, max_per_table: int
) -> list[int] | None:
    if table_count <= 0 or min_per_table < 0 or max_per_table < min_per_table:
        return None

    min_total = table_count * min_per_table
    max_total = table_count * max_per_table
    if total_students < min_total or total_students > max_total:
        return None

    sizes = [min_per_table for _ in range(table_count)]
    remaining = total_students - min_total
    idx = 0
    while remaining > 0:
        if sizes[idx] < max_per_table:
            sizes[idx] += 1
            remaining -= 1
        idx = (idx + 1) % table_count
    return sizes


def _can_share_table(skill_a: str, skill_b: str) -> bool:
    return not _is_forbidden_skill_pair(skill_a, skill_b)


def _members_accept_skill(members: list[Student], skill_level: str) -> bool:
    return all(_can_share_table(skill_level, member.skill_level) for member in members)


def _members_accept_skill_with_extra(
    members: list[Student], extra_skill_level: str, skill_level: str
) -> bool:
    if not _can_share_table(skill_level, extra_skill_level):
        return False
    return _members_accept_skill(members, skill_level)


def _has_future_capacity_after_placement(
    *,
    tables: list[list[Student]],
    target_sizes: list[int],
    remaining_skill_counts: Counter[str],
    place_table_idx: int,
    place_student_skill: str,
) -> bool:
    # Simulate consuming current student.
    rest_counts = Counter(remaining_skill_counts)
    rest_counts[place_student_skill] -= 1
    if rest_counts[place_student_skill] <= 0:
        rest_counts.pop(place_student_skill, None)

    for skill_level, need_count in rest_counts.items():
        if need_count <= 0:
            continue

        available_slots = 0
        for table_idx, members in enumerate(tables):
            current_size = len(members)
            if table_idx == place_table_idx:
                current_size += 1
                if current_size >= target_sizes[table_idx]:
                    continue
                if _members_accept_skill_with_extra(members, place_student_skill, skill_level):
                    available_slots += target_sizes[table_idx] - current_size
            else:
                if current_size >= target_sizes[table_idx]:
                    continue
                if _members_accept_skill(members, skill_level):
                    available_slots += target_sizes[table_idx] - current_size

        if need_count > available_slots:
            return False

    return True


def _capacity_fill_rows(students: list[Student], config: SeatingConfig, rng: random.Random) -> list[dict[str, Any]]:
    target_sizes = _bounded_balanced_table_sizes(
        total_students=len(students),
        table_count=config.table_count,
        min_per_table=config.min_per_table,
        max_per_table=config.max_per_table,
    )
    if target_sizes is None:
        raise ValueError(
            "人数制約を満たせません。テーブル数・最小人数・最大人数を見直してください。"
        )

    # Hard-skill-constraint fallback:
    # keep forbidden pairs out while trying to fill all seats.
    tables: list[list[Student]] = [[] for _ in range(config.table_count)]
    table_company_counts: list[defaultdict[str, int]] = [defaultdict(int) for _ in range(config.table_count)]
    remaining = students[:]
    rng.shuffle(remaining)

    while remaining:
        # Most-constrained-first: choose student with fewest compatible tables.
        best_idx: int | None = None
        best_candidates: list[tuple[int, int]] = []

        for idx, student in enumerate(remaining):
            student_company_key = _normalize_company_key(student.company)
            candidate_tables: list[tuple[int, int, int]] = []
            for table_idx, members in enumerate(tables):
                if len(members) >= target_sizes[table_idx]:
                    continue
                if any(_is_forbidden_skill_pair(student.skill_level, m.skill_level) for m in members):
                    continue
                soft_avoid_count = sum(
                    1 for m in members if _is_soft_avoid_pair(student.skill_level, m.skill_level)
                )
                same_company_count = table_company_counts[table_idx][student_company_key]
                candidate_tables.append((table_idx, soft_avoid_count, same_company_count))

            if best_idx is None or len(candidate_tables) < len(best_candidates):
                best_idx = idx
                best_candidates = candidate_tables
                if len(best_candidates) == 0:
                    break

        if best_idx is None or not best_candidates:
            raise ValueError(
                "スキル同席制約のため配席できません。テーブル数や min/max 設定を見直してください。"
            )

        student = remaining.pop(best_idx)
        student_company_key = _normalize_company_key(student.company)
        # Prefer no 中×低 conflicts, then fewer same-company members, then larger remaining capacity.
        min_soft_conflict = min(soft for _, soft, _ in best_candidates)
        filtered_candidates = [item for item in best_candidates if item[1] == min_soft_conflict]
        min_same_company = min(same_company for _, _, same_company in filtered_candidates)
        filtered_candidates = [
            item for item in filtered_candidates if item[2] == min_same_company
        ]
        filtered_candidates.sort(
            key=lambda t: (
                -(target_sizes[t[0]] - len(tables[t[0]])),
                abs((len(tables[t[0]]) + 1) - target_sizes[t[0]]),
                rng.random(),
            )
        )
        chosen_table = filtered_candidates[0][0]
        tables[chosen_table].append(student)
        table_company_counts[chosen_table][student_company_key] += 1
    return _students_to_rows(tables)


def _students_to_rows(tables: list[list[Student]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for table_idx, members in enumerate(tables, start=1):
        for student in members:
            rows.append(
                {
                    "table_no": table_idx,
                    "student_id": student.id,
                    "name": student.name,
                    "name_kana": student.name_kana,
                    "company": student.company,
                    "skill_level": student.skill_level,
                }
            )
    rows.sort(key=lambda r: (int(r["table_no"]), int(r["student_id"])))
    return rows


def _dominant_table_skill(members: list[dict[str, Any]]) -> str:
    if not members:
        return ""
    counts = Counter(str(m.get("skill_level", "")).strip() for m in members)
    ranked = sorted(
        counts.items(),
        key=lambda item: (-item[1], _SKILL_PRIORITY.get(item[0], len(_SKILL_PRIORITY)), item[0]),
    )
    return ranked[0][0] if ranked else ""


def _renumber_tables_by_skill_priority(
    rows: list[dict[str, Any]], table_count: int
) -> list[dict[str, Any]]:
    grouped: dict[int, list[dict[str, Any]]] = {table_no: [] for table_no in range(1, table_count + 1)}
    for row in rows:
        table_no = int(row["table_no"])
        if table_no in grouped:
            grouped[table_no].append(row)

    source_table_order = sorted(
        grouped.keys(),
        key=lambda table_no: (
            _SKILL_PRIORITY.get(
                _dominant_table_skill(grouped[table_no]),
                len(_SKILL_PRIORITY),
            ),
            table_no,
        ),
    )

    remap: dict[int, int] = {
        old_table_no: new_table_no for new_table_no, old_table_no in enumerate(source_table_order, start=1)
    }

    renumbered: list[dict[str, Any]] = []
    for row in rows:
        old_table_no = int(row["table_no"])
        new_row = dict(row)
        new_row["table_no"] = remap.get(old_table_no, old_table_no)
        renumbered.append(new_row)

    renumbered.sort(key=lambda r: (int(r["table_no"]), int(r["student_id"])))
    return renumbered


def _soft_conflict_pairs_from_skills(skills: list[str]) -> int:
    normal_count = sum(1 for s in skills if s == "並")
    yabai_count = sum(1 for s in skills if s == "ヤバい")
    return normal_count * yabai_count


def _soft_avoid_pair_total(rows: list[dict[str, Any]], table_count: int) -> int:
    grouped_skills: dict[int, list[str]] = {table_no: [] for table_no in range(1, table_count + 1)}
    for row in rows:
        table_no = int(row["table_no"])
        if table_no in grouped_skills:
            grouped_skills[table_no].append(str(row.get("skill_level", "")).strip())
    return sum(_soft_conflict_pairs_from_skills(skills) for skills in grouped_skills.values())


def _company_collision_count_from_companies(companies: list[str]) -> int:
    normalized_counter = Counter(_normalize_company_key(company) for company in companies if company)
    return sum(count - 1 for count in normalized_counter.values() if count > 1)


def _company_collision_total(rows: list[dict[str, Any]], table_count: int) -> int:
    grouped_companies: dict[int, list[str]] = {table_no: [] for table_no in range(1, table_count + 1)}
    for row in rows:
        table_no = int(row["table_no"])
        if table_no in grouped_companies:
            grouped_companies[table_no].append(str(row.get("company", "")).strip())
    return sum(
        _company_collision_count_from_companies(companies)
        for companies in grouped_companies.values()
    )


def _has_hard_conflict_in_skills(skills: list[str]) -> bool:
    for i in range(len(skills)):
        for j in range(i + 1, len(skills)):
            if _is_forbidden_skill_pair(skills[i], skills[j]):
                return True
    return False


def _optimize_soft_avoid_pairs(rows: list[dict[str, Any]], table_count: int) -> list[dict[str, Any]]:
    grouped: dict[int, list[dict[str, Any]]] = {table_no: [] for table_no in range(1, table_count + 1)}
    working_rows = [dict(row) for row in rows]
    for row in working_rows:
        grouped[int(row["table_no"])].append(row)

    max_passes = 400
    for _ in range(max_passes):
        improved = False
        for left_table in range(1, table_count + 1):
            left_members = grouped[left_table]
            if not left_members:
                continue
            left_skills = [str(m["skill_level"]) for m in left_members]
            left_soft_before = _soft_conflict_pairs_from_skills(left_skills)

            for right_table in range(left_table + 1, table_count + 1):
                right_members = grouped[right_table]
                if not right_members:
                    continue
                right_skills = [str(m["skill_level"]) for m in right_members]
                right_soft_before = _soft_conflict_pairs_from_skills(right_skills)
                before_total = left_soft_before + right_soft_before

                for li, left_student in enumerate(left_members):
                    left_skill = str(left_student["skill_level"])
                    for ri, right_student in enumerate(right_members):
                        right_skill = str(right_student["skill_level"])
                        if left_skill == right_skill:
                            continue

                        next_left_skills = left_skills[:]
                        next_right_skills = right_skills[:]
                        next_left_skills[li] = right_skill
                        next_right_skills[ri] = left_skill

                        if _has_hard_conflict_in_skills(next_left_skills):
                            continue
                        if _has_hard_conflict_in_skills(next_right_skills):
                            continue

                        after_total = (
                            _soft_conflict_pairs_from_skills(next_left_skills)
                            + _soft_conflict_pairs_from_skills(next_right_skills)
                        )
                        if after_total >= before_total:
                            continue

                        # Apply improving swap.
                        left_members[li], right_members[ri] = right_student, left_student
                        left_members[li]["table_no"] = left_table
                        right_members[ri]["table_no"] = right_table
                        improved = True
                        break
                    if improved:
                        break
                if improved:
                    break
            if improved:
                break
        if not improved:
            break

    optimized: list[dict[str, Any]] = []
    for table_no in range(1, table_count + 1):
        optimized.extend(grouped[table_no])
    optimized.sort(key=lambda r: (int(r["table_no"]), int(r["student_id"])))
    return optimized


def _optimize_company_collisions(rows: list[dict[str, Any]], table_count: int) -> list[dict[str, Any]]:
    grouped: dict[int, list[dict[str, Any]]] = {table_no: [] for table_no in range(1, table_count + 1)}
    working_rows = [dict(row) for row in rows]
    for row in working_rows:
        grouped[int(row["table_no"])].append(row)

    max_passes = 500
    for _ in range(max_passes):
        improved = False
        for left_table in range(1, table_count + 1):
            left_members = grouped[left_table]
            if not left_members:
                continue
            left_skills = [str(m["skill_level"]) for m in left_members]
            left_companies = [str(m.get("company", "")).strip() for m in left_members]

            for right_table in range(left_table + 1, table_count + 1):
                right_members = grouped[right_table]
                if not right_members:
                    continue
                right_skills = [str(m["skill_level"]) for m in right_members]
                right_companies = [str(m.get("company", "")).strip() for m in right_members]

                before_company_total = (
                    _company_collision_count_from_companies(left_companies)
                    + _company_collision_count_from_companies(right_companies)
                )
                before_soft_total = (
                    _soft_conflict_pairs_from_skills(left_skills)
                    + _soft_conflict_pairs_from_skills(right_skills)
                )

                for li, left_student in enumerate(left_members):
                    left_skill = str(left_student["skill_level"])
                    left_company = str(left_student.get("company", "")).strip()
                    left_company_key = _normalize_company_key(left_company)
                    for ri, right_student in enumerate(right_members):
                        right_skill = str(right_student["skill_level"])
                        right_company = str(right_student.get("company", "")).strip()
                        right_company_key = _normalize_company_key(right_company)
                        if left_company_key == right_company_key:
                            continue

                        next_left_skills = left_skills[:]
                        next_right_skills = right_skills[:]
                        next_left_skills[li] = right_skill
                        next_right_skills[ri] = left_skill
                        if _has_hard_conflict_in_skills(next_left_skills):
                            continue
                        if _has_hard_conflict_in_skills(next_right_skills):
                            continue

                        next_left_companies = left_companies[:]
                        next_right_companies = right_companies[:]
                        next_left_companies[li] = right_company
                        next_right_companies[ri] = left_company

                        after_company_total = (
                            _company_collision_count_from_companies(next_left_companies)
                            + _company_collision_count_from_companies(next_right_companies)
                        )
                        if after_company_total >= before_company_total:
                            continue

                        after_soft_total = (
                            _soft_conflict_pairs_from_skills(next_left_skills)
                            + _soft_conflict_pairs_from_skills(next_right_skills)
                        )
                        if after_soft_total > before_soft_total + 2:
                            continue

                        left_members[li], right_members[ri] = right_student, left_student
                        left_members[li]["table_no"] = left_table
                        right_members[ri]["table_no"] = right_table
                        improved = True
                        break
                    if improved:
                        break
                if improved:
                    break
            if improved:
                break
        if not improved:
            break

    optimized: list[dict[str, Any]] = []
    for table_no in range(1, table_count + 1):
        optimized.extend(grouped[table_no])
    optimized.sort(key=lambda r: (int(r["table_no"]), int(r["student_id"])))
    return optimized


def optimize_soft_skill_conflicts(rows: list[dict[str, Any]], table_count: int) -> list[dict[str, Any]]:
    return _optimize_soft_avoid_pairs(rows, table_count)


def count_soft_avoid_pairs(rows: list[dict[str, Any]], table_count: int) -> int:
    return int(_soft_avoid_pair_total(rows, table_count))


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
    expected_sizes = _bounded_balanced_table_sizes(
        total_students=total_students,
        table_count=config.table_count,
        min_per_table=config.min_per_table,
        max_per_table=config.max_per_table,
    )
    if expected_sizes is None:
        expected_sizes = [0 for _ in range(config.table_count)]

    company_collision_total = 0
    previous_pair_hits = 0
    skill_mixing_total = 0.0
    exact_skill_mixing_total = 0.0
    forbidden_skill_pair_total = 0
    soft_avoid_pair_total = 0
    size_deviation_total = 0.0
    current_pairs: set[tuple[int, int]] = set()

    table_metrics: list[dict[str, Any]] = []

    for table_no in range(1, config.table_count + 1):
        members = table_members[table_no]
        company_counter = Counter(
            _normalize_company_key(str(m.get("company", ""))) for m in members
        )
        company_display_by_key: dict[str, str] = {}
        for member in members:
            raw_company = str(member.get("company", "")).strip()
            key = _normalize_company_key(raw_company)
            if key and key not in company_display_by_key:
                company_display_by_key[key] = raw_company
        skill_counter = Counter(m["skill_level"] for m in members)
        skill_group_counter = Counter(_skill_group(str(m["skill_level"])) for m in members)
        duplicate_companies = {
            company_display_by_key.get(company_key, company_key): count
            for company_key, count in company_counter.items()
            if count > 1
        }
        company_collisions = sum(count - 1 for count in company_counter.values() if count > 1)
        company_collision_total += company_collisions

        for i in range(len(members)):
            sid_i = int(members[i]["student_id"])
            skill_i = str(members[i]["skill_level"])
            for j in range(i + 1, len(members)):
                sid_j = int(members[j]["student_id"])
                skill_j = str(members[j]["skill_level"])
                pair = _pair_key(sid_i, sid_j)
                current_pairs.add(pair)
                if pair in previous_pairs:
                    previous_pair_hits += 1
                if _is_forbidden_skill_pair(skill_i, skill_j):
                    forbidden_skill_pair_total += 1

        total_pairs = len(members) * (len(members) - 1) // 2
        soft_avoid_pair_total += _soft_conflict_pairs_from_skills(
            [str(m["skill_level"]) for m in members]
        )
        same_group_pairs = sum(
            count * (count - 1) // 2 for count in skill_group_counter.values()
        )
        skill_mixing_pairs = total_pairs - same_group_pairs
        same_skill_pairs = sum(count * (count - 1) // 2 for count in skill_counter.values())
        exact_skill_mixing_pairs = total_pairs - same_skill_pairs
        skill_mixing_total += skill_mixing_pairs
        exact_skill_mixing_total += exact_skill_mixing_pairs

        target_size = expected_sizes[table_no - 1]
        size_deviation_total += abs(len(members) - target_size)

        table_metrics.append(
            {
                "table_no": table_no,
                "size": len(members),
                "skill_counts": {skill: skill_counter.get(skill, 0) for skill in SKILL_LEVELS},
                "skill_group_counts": dict(skill_group_counter),
                "skill_mixing_pairs": int(skill_mixing_pairs),
                "exact_skill_mixing_pairs": int(exact_skill_mixing_pairs),
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
        + (skill_mixing_total + 0.8 * exact_skill_mixing_total) * config.skill_weight
        + size_deviation_total * config.size_weight
        + forbidden_skill_pair_total * config.skill_weight * 1000.0
    )

    metrics = {
        "company_collision_total": company_collision_total,
        "previous_pair_hits": previous_pair_hits,
        "skill_mixing_total": round(skill_mixing_total, 3),
        "exact_skill_mixing_total": round(exact_skill_mixing_total, 3),
        "forbidden_skill_pair_total": forbidden_skill_pair_total,
        "soft_avoid_pair_total": int(soft_avoid_pair_total),
        # Keep the old key so existing UI/session data can still read it.
        "skill_deviation_total": round(skill_mixing_total, 3),
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
    enforce_future_feasibility: bool,
) -> tuple[list[dict[str, Any]], float]:
    table_count = config.table_count
    max_per_table = config.max_per_table
    total_students = len(students)

    tables: list[list[Student]] = [[] for _ in range(table_count)]
    table_company_counts: list[defaultdict[str, int]] = [defaultdict(int) for _ in range(table_count)]
    table_skill_group_counts: list[defaultdict[str, int]] = [defaultdict(int) for _ in range(table_count)]
    table_skill_level_counts: list[defaultdict[str, int]] = [
        defaultdict(int) for _ in range(table_count)
    ]

    target_sizes = _bounded_balanced_table_sizes(
        total_students=total_students,
        table_count=table_count,
        min_per_table=config.min_per_table,
        max_per_table=max_per_table,
    )
    if target_sizes is None:
        return [], math.inf

    skill_group_counts = Counter(_skill_group(s.skill_level) for s in students)
    skill_counts = Counter(s.skill_level for s in students)
    company_counts = Counter(_normalize_company_key(s.company) for s in students)

    order = students[:]
    rng.shuffle(order)
    order.sort(
        key=lambda s: (
            _SKILL_PRIORITY.get(s.skill_level, len(_SKILL_PRIORITY)),  # 高 -> 中 -> やや低 -> 低
            skill_group_counts.get(_skill_group(s.skill_level), 0),  # レアなグループを先に置く
            skill_counts.get(s.skill_level, 0),  # 同グループ内のレア度
            -company_counts.get(_normalize_company_key(s.company), 0),  # 人数が多い会社を先に分散
            rng.random(),
        )
    )
    remaining_skill_counts = Counter(s.skill_level for s in order)

    for student in order:
        student_group = _skill_group(student.skill_level)
        student_company_key = _normalize_company_key(student.company)
        candidates: list[tuple[float, int, int]] = []
        for table_idx in range(table_count):
            members = tables[table_idx]
            if len(members) >= target_sizes[table_idx]:
                continue

            forbidden_with_members = sum(
                1
                for other in members
                if _is_forbidden_skill_pair(student.skill_level, other.skill_level)
            )
            if forbidden_with_members > 0:
                # Hard constraint: forbidden pairs cannot share a table.
                continue

            if enforce_future_feasibility:
                if not _has_future_capacity_after_placement(
                    tables=tables,
                    target_sizes=target_sizes,
                    remaining_skill_counts=remaining_skill_counts,
                    place_table_idx=table_idx,
                    place_student_skill=student.skill_level,
                ):
                    continue

            same_company_count = table_company_counts[table_idx][student_company_key]
            company_penalty = (
                ((same_company_count + 1) ** 2 - 1) * config.company_weight
            )

            prev_hits = 0
            for other in members:
                if _pair_key(student.id, other.id) in previous_pairs:
                    prev_hits += 1
            previous_penalty = prev_hits * config.previous_weight

            different_group_count = sum(
                count
                for group, count in table_skill_group_counts[table_idx].items()
                if group != student_group
            )
            same_group_count = table_skill_group_counts[table_idx][student_group]
            same_skill_count = table_skill_level_counts[table_idx][student.skill_level]
            adjacent_skill_count = 0
            far_skill_count = 0
            soft_avoid_count = 0
            for other in members:
                other_skill = other.skill_level
                if _is_soft_avoid_pair(student.skill_level, other_skill):
                    soft_avoid_count += 1
                dist = abs(
                    _SKILL_PRIORITY.get(student.skill_level, 99)
                    - _SKILL_PRIORITY.get(other_skill, 99)
                )
                if dist == 1:
                    adjacent_skill_count += 1
                elif dist >= 2:
                    far_skill_count += 1

            # Same-skill is best. Adjacent (高-中 / 中-やや低 / やや低-低) is next best.
            # Far pairs, especially 中×低, are strongly discouraged.
            skill_penalty = (
                0.8 * adjacent_skill_count
                + 7.5 * far_skill_count
                + 12.0 * soft_avoid_count
                - 1.2 * same_skill_count
                - 0.45 * same_group_count
            ) * config.skill_weight
            if adjacent_skill_count > 0:
                skill_penalty += (adjacent_skill_count**2) * 0.10 * config.skill_weight
            if far_skill_count > 0:
                skill_penalty += (far_skill_count**2) * 0.80 * config.skill_weight

            target_size = target_sizes[table_idx]
            before_size = abs(len(members) - target_size)
            after_size = abs(len(members) + 1 - target_size)
            size_penalty = (after_size - before_size) * config.size_weight

            jitter = rng.random() * config.randomness_weight
            score = company_penalty + previous_penalty + skill_penalty + size_penalty + jitter
            candidates.append((score, table_idx, soft_avoid_count))

        if not candidates:
            return [], math.inf

        min_soft_avoid = min(c[2] for c in candidates)
        if min_soft_avoid == 0:
            candidates = [c for c in candidates if c[2] == 0]

        candidates.sort(key=lambda x: x[0])
        _, chosen_table, _ = candidates[0]

        tables[chosen_table].append(student)
        table_company_counts[chosen_table][student_company_key] += 1
        table_skill_group_counts[chosen_table][student_group] += 1
        table_skill_level_counts[chosen_table][student.skill_level] += 1
        remaining_skill_counts[student.skill_level] -= 1
        if remaining_skill_counts[student.skill_level] <= 0:
            remaining_skill_counts.pop(student.skill_level, None)

    rows = _students_to_rows(tables)
    score, _ = _evaluate_assignment(rows, previous_pairs, previous_table_map={}, config=config)
    return rows, score


def generate_best_assignment(
    students: list[Student],
    previous_pairs: set[tuple[int, int]],
    previous_table_map: dict[int, int],
    config: SeatingConfig,
) -> dict[str, Any]:
    students = _dedupe_students_for_assignment(students)

    if config.min_per_table > config.max_per_table:
        raise ValueError(
            f"最小人数と最大人数の設定が不正です。min={config.min_per_table}, max={config.max_per_table}"
        )

    min_required = config.table_count * config.min_per_table
    capacity = config.table_count * config.max_per_table
    if len(students) < min_required:
        raise ValueError(
            "受講生数が最小必要人数を下回っています。"
            f" students={len(students)}, min_required={min_required}"
        )
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
                "skill_mixing_total": 0.0,
                "skill_deviation_total": 0.0,
                "forbidden_skill_pair_total": 0,
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
    used_mode = "strict"

    search_modes: list[tuple[str, bool, int]] = [
        ("strict", True, max(1, config.attempts)),
        ("strict_relaxed", False, max(1, config.attempts)),
    ]

    for mode_name, enforce_future_feasibility, attempts in search_modes:
        mode_best_rows: list[dict[str, Any]] = []
        mode_best_score = math.inf
        for _ in range(attempts):
            attempt_rng = random.Random(base_rng.randint(1, 10**9))
            rows, rough_score = _single_attempt(
                students=students,
                previous_pairs=previous_pairs,
                config=config,
                rng=attempt_rng,
                enforce_future_feasibility=enforce_future_feasibility,
            )
            if not rows:
                continue
            if rough_score < mode_best_score:
                mode_best_rows = rows
                mode_best_score = rough_score
        if mode_best_rows:
            best_rows = mode_best_rows
            best_score = mode_best_score
            used_mode = mode_name
            break

    if not best_rows:
        # Hard-constraint fallback: still disallow forbidden skill pairs.
        best_rows = _capacity_fill_rows(students, config, random.Random(base_rng.randint(1, 10**9)))
        used_mode = "hard_constraint_fallback"

    best_rows = _optimize_company_collisions(best_rows, config.table_count)
    best_rows = _optimize_soft_avoid_pairs(best_rows, config.table_count)
    best_rows = _optimize_company_collisions(best_rows, config.table_count)
    best_rows = _renumber_tables_by_skill_priority(best_rows, config.table_count)
    _validate_no_duplicate_assignment_rows(best_rows, expected_total=len(students))
    final_score, metrics = _evaluate_assignment(best_rows, previous_pairs, previous_table_map, config)
    metrics["assignment_mode"] = used_mode
    metrics["fallback_used"] = used_mode != "strict"
    return {"rows": best_rows, "score": final_score, "metrics": metrics}


def validate_manual_rows(
    rows: list[dict[str, Any]], table_count: int, min_per_table: int, max_per_table: int
) -> list[str]:
    errors: list[str] = []
    if any(int(row["table_no"]) < 1 or int(row["table_no"]) > table_count for row in rows):
        errors.append(f"table_no は 1〜{table_count} の範囲で入力してください。")

    duplicate_ids = _find_duplicate_student_ids(rows)
    if duplicate_ids:
        preview = ", ".join(str(student_id) for student_id in duplicate_ids[:10])
        suffix = " ..." if len(duplicate_ids) > 10 else ""
        errors.append(
            "同じ受講生IDが複数テーブルに重複配置されています。"
            f" student_id={preview}{suffix}"
        )

    duplicate_identities = _find_duplicate_row_identities(rows)
    if duplicate_identities:
        preview = " / ".join(duplicate_identities[:6])
        suffix = " ..." if len(duplicate_identities) > 6 else ""
        errors.append(
            "同一受講生（氏名+会社）の重複配置があります。"
            f" {preview}{suffix}"
        )

    table_sizes = Counter(int(row["table_no"]) for row in rows)
    under_min = [table_no for table_no in range(1, table_count + 1) if table_sizes.get(table_no, 0) < min_per_table]
    if under_min:
        under_min_str = ", ".join(str(no) for no in under_min)
        errors.append(
            f"テーブル下限({min_per_table}名)を下回っているテーブルがあります: {under_min_str}"
        )

    over_capacity = [table_no for table_no, size in table_sizes.items() if size > max_per_table]
    if over_capacity:
        over_capacity_str = ", ".join(str(no) for no in sorted(over_capacity))
        errors.append(
            f"テーブル上限({max_per_table}名)を超えているテーブルがあります: {over_capacity_str}"
        )

    rows_by_table: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        rows_by_table[int(row["table_no"])].append(row)

    skill_conflict_tables: list[int] = []
    for table_no, members in rows_by_table.items():
        conflict_found = False
        for i in range(len(members)):
            for j in range(i + 1, len(members)):
                if _is_forbidden_skill_pair(
                    str(members[i]["skill_level"]), str(members[j]["skill_level"])
                ):
                    conflict_found = True
                    break
            if conflict_found:
                break
        if conflict_found:
            skill_conflict_tables.append(table_no)

    if skill_conflict_tables:
        table_str = ", ".join(str(no) for no in sorted(skill_conflict_tables))
        errors.append(
            "スキル制約に違反する同席があります。"
            f"（高×やや低 / 高×低） テーブル: {table_str}"
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
