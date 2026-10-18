"""Independent attainment calculator for uniform, complete score snapshots.

Standard library only. Does not access Excel, generate scores, or call an AI.
See references/excel.md for the input and observed-result contracts.
"""
import argparse
import csv
import hashlib
import json
import sys
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
from pathlib import Path


def number(value, label):
    if isinstance(value, bool) or value is None or str(value).strip() == "":
        raise ValueError(f"{label}: missing or nonnumeric value")
    try:
        result = Decimal(str(value))
    except InvalidOperation:
        raise ValueError(f"{label}: invalid number {value!r}") from None
    if not result.is_finite():
        raise ValueError(f"{label}: nonfinite number")
    return result


def calculate(rows, config):
    identity = config["identity_columns"]
    if not identity or len(set(identity)) != len(identity):
        raise ValueError("identity_columns must be nonempty and distinct")
    decimals = config["decimals"]
    if type(decimals) is not int or not 0 <= decimals <= 12:
        raise ValueError("decimals must be an integer in 0..12")
    count = config.get("expected_count")
    if count is not None and (type(count) is not int or count <= 0):
        raise ValueError("expected_count must be a positive integer")
    if not rows or (count is not None and count != len(rows)):
        raise ValueError("empty snapshot or unexpected student count")
    goals = config["objectives"]
    goal_ids = [goal["id"] for goal in goals]
    if not goal_ids or any(not isinstance(k, str) or not k for k in goal_ids) or len(set(goal_ids)) != len(goal_ids):
        raise ValueError("objective IDs must be nonempty and distinct")
    parsed = {}
    maxima = {column: number(value, f"maximum {column}") for column, value in config.get("column_maxima", {}).items()}
    if any(value <= 0 for value in maxima.values()):
        raise ValueError("column_maxima must be positive")
    for goal in goals:
        items, used = [], set()
        for item in goal["items"]:
            columns = item["columns"]
            if not columns or len(set(columns)) != len(columns) or used.intersection(columns):
                raise ValueError(f"{goal['id']}: missing or repeated score columns")
            used.update(columns)
            full = number(item["full_score"], "full_score")
            weight = number(item["contribution"], "contribution")
            if full <= 0 or weight < 0:
                raise ValueError("full_score must be positive; contribution must be nonnegative")
            items.append((columns, full, weight))
        if not items or sum(w for _, _, w in items) <= 0:
            raise ValueError(f"{goal['id']}: zero total contribution")
        parsed[goal["id"]] = items
    with localcontext() as ctx:
        ctx.prec = 50
        unit = Decimal(1).scaleb(-decimals)
        totals = {key: Decimal(0) for key in goal_ids}
        students, seen = [], set()
        for index, row in enumerate(rows, start=1):
            key = tuple(str(row[col]) for col in identity)
            if any(not k.strip() for k in key) or key in seen:
                raise ValueError(f"record {index}: blank or duplicate student key")
            seen.add(key)
            attainment = {}
            for goal_id, items in parsed.items():
                numerator = Decimal(0)
                denominator = sum(w for _, _, w in items)
                for columns, full, weight in items:
                    parts = [number(row[col], f"record {index}, {col}") for col in columns]
                    if any(col in maxima and part > maxima[col] for col, part in zip(columns, parts)):
                        raise ValueError(f"record {index}, {columns}: component exceeds column maximum")
                    score = sum(parts)
                    if any(part < 0 for part in parts) or score > full:
                        raise ValueError(f"record {index}, {columns}: score outside 0..{full}")
                    numerator += score / full * weight
                exact = numerator / denominator
                totals[goal_id] += exact
                attainment[goal_id] = {"exact": str(exact), "rounded": str(exact.quantize(unit, rounding=ROUND_HALF_UP))}
            students.append({"key": list(key), "attainment": attainment})
        aggregate = {}
        for goal_id, total in totals.items():
            exact = total / len(rows)
            aggregate[goal_id] = {"exact": str(exact), "rounded": str(exact.quantize(unit, rounding=ROUND_HALF_UP))}
        return {"student_count": len(rows), "students": students, "aggregate": aggregate}


def verify(result, observed):
    if [s["key"] for s in result["students"]] != [s["key"] for s in observed["students"]]:
        raise ValueError("observed student keys, count, or order differ")
    for index, (expected, actual) in enumerate(zip(result["students"], observed["students"]), start=1):
        if set(expected["attainment"]) != set(actual["attainment"]):
            raise ValueError(f"record {index}: objective set differs")
        for goal, values in expected["attainment"].items():
            if number(values["rounded"], "expected") != number(actual["attainment"][goal]["rounded"], "observed"):
                raise ValueError(f"record {index}, {goal}: rounded attainment differs")
    if set(result["aggregate"]) != set(observed["aggregate"]):
        raise ValueError("observed aggregate objective set differs")
    with localcontext() as ctx:
        ctx.prec = 50
        for goal, values in result["aggregate"].items():
            actual = observed["aggregate"][goal]
            if abs(number(values["exact"], "expected") - number(actual["exact"], "observed")) > Decimal("1e-9"):
                raise ValueError(f"aggregate {goal}: exact attainment differs")
            if number(values["rounded"], "expected") != number(actual["rounded"], "observed"):
                raise ValueError(f"aggregate {goal}: rounded attainment differs")


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--csv", required=True, type=Path)
    parser.add_argument("--config", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--observed", type=Path)
    args = parser.parse_args()
    inputs = [args.csv, args.config] + ([args.observed] if args.observed else [])
    if args.output.resolve() in [p.resolve() for p in inputs]:
        parser.error("output must differ from input paths")
    try:
        config = json.loads(args.config.read_text(encoding="utf-8-sig"))
        with args.csv.open(encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            if not reader.fieldnames or len(set(reader.fieldnames)) != len(reader.fieldnames):
                raise ValueError("missing or duplicate CSV headers")
            rows = list(reader)
            if any(None in row or any(value is None for value in row.values()) for row in rows):
                raise ValueError("CSV row width differs from header width")
        result = calculate(rows, config)
        if args.observed:
            verify(result, json.loads(args.observed.read_text(encoding="utf-8-sig")))
        result["verification"] = "excel_numeric_results_matched" if args.observed else "computed_only"
        result["input_sha256"] = {str(path): digest(path) for path in inputs}
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({"students": result["student_count"], "aggregate": result["aggregate"], "verification": result["verification"], "output": str(args.output)}, ensure_ascii=False))
        return 0
    except (ValueError, KeyError, TypeError, OSError, csv.Error) as error:
        print(f"Attainment check failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
