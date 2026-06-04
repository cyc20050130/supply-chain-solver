from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

import pandas as pd

from solve import cost_score, deviation_score, satisfaction_score


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


SCORE_POINTS = {
    "采购": {
        "单位采购成本": 60.0,
        "生产满足率": 40.0,
    },
    "生产": {
        "单位物流成本": 60.0,
        "市场满足率": 40.0,
    },
    "销售": {
        "预测偏差率": 20.0,
        "单位物流成本": 50.0,
        "市场满足率": 30.0,
    },
    "综合": {
        "预测偏差率": 10.0,
        "单位物流成本": 20.0,
        "单位采购成本": 20.0,
        "生产满足率": 15.0,
        "市场满足率": 35.0,
    },
}


def number(text: Any) -> float:
    cleaned = re.sub(r"[^0-9.\-]", "", "" if text is None else str(text))
    if cleaned in ("", ".", "-", "-."):
        return 0.0
    return float(cleaned)


def parse_ratio_or_money(text: Any) -> float:
    raw = "" if text is None else str(text)
    value = number(raw)
    if "%" in raw or "％" in raw:
        return value / 100.0
    return value


def parse_reported_score(text: Any) -> float:
    raw = "" if text is None else str(text)
    match = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*分", raw)
    if match:
        return float(match.group(1))
    return number(raw)


def parse_score_page(path: Path) -> dict[str, Any]:
    tables = pd.read_html(str(path), encoding="utf-8")
    if not tables:
        raise ValueError(f"成绩页没有可解析表格: {path}")
    table = tables[0].iloc[:, :4]
    known_items = set().union(*(items.keys() for items in SCORE_POINTS.values()))
    rows = []
    for _, row in table.iterrows():
        item = str(row.iloc[0]).strip()
        if item not in known_items:
            continue
        target = parse_ratio_or_money(row.iloc[1])
        actual = parse_ratio_or_money(row.iloc[2])
        reported_score = parse_reported_score(row.iloc[3])
        rows.append(
            {
                "item": item,
                "target": target,
                "actual": actual,
                "reported_score": reported_score,
                "reported_text": str(row.iloc[3]),
            }
        )
    problem_rows = []
    for _, row in table.iterrows():
        item = str(row.iloc[0]).strip()
        if "断货" in item or "断料" in item:
            problem_rows.append({"item": item, "amount_wan": number(row.iloc[1]), "note": str(row.iloc[2])})
    return {"path": str(path), "rows": rows, "problem_rows": problem_rows}


def replay_rows(rows: list[dict[str, Any]], mode: str) -> dict[str, Any]:
    replayed = []
    total = 0.0
    for row in rows:
        item = row["item"]
        target = row["target"]
        actual = row["actual"]
        points = SCORE_POINTS[mode][item]
        if "成本" in item:
            calc = cost_score(actual, target, points)
        elif item == "预测偏差率":
            calc = deviation_score(actual, target, points)
        else:
            calc = satisfaction_score(actual, points)
        total += calc
        replayed.append({**row, "calculated_score": round(calc, 4), "diff": round(calc - row.get("reported_score", calc), 4), "max": points})
    return {"mode": mode, "rows": replayed, "total": round(total, 4)}


def built_in_regression() -> dict[str, Any]:
    procurement_rows = [
        {"item": "单位采购成本", "target": 415.0, "actual": 423.71, "reported_score": 53.7},
        {"item": "生产满足率", "target": 1.0, "actual": 0.9842, "reported_score": 39.37},
    ]
    comprehensive_rows = [
        {"item": "预测偏差率", "target": 0.05, "actual": 0.1612, "reported_score": 0.0},
        {"item": "单位物流成本", "target": 110.0, "actual": 107.90, "reported_score": 20.0},
        {"item": "单位采购成本", "target": 1600.0, "actual": 1809.29, "reported_score": 6.92},
        {"item": "生产满足率", "target": 1.0, "actual": 0.8198, "reported_score": 12.3},
        {"item": "市场满足率", "target": 1.0, "actual": 0.66, "reported_score": 23.1},
    ]
    sales_98_rows = [
        {"item": "预测偏差率", "target": 0.05, "actual": 0.0, "reported_score": 20.0},
        {"item": "单位物流成本", "target": 50.0, "actual": 49.15, "reported_score": 50.0},
        {"item": "市场满足率", "target": 1.0, "actual": 0.9486, "reported_score": 28.46},
    ]
    sales_82_rows = [
        {"item": "预测偏差率", "target": 0.05, "actual": 0.04, "reported_score": 4.0},
        {"item": "单位物流成本", "target": 50.0, "actual": 49.37, "reported_score": 50.0},
        {"item": "市场满足率", "target": 1.0, "actual": 0.9417, "reported_score": 28.25},
    ]
    production_rows = [
        {"item": "单位物流成本", "target": 190.0, "actual": 163.85, "reported_score": 60.0},
        {"item": "市场满足率", "target": 1.0, "actual": 0.991, "reported_score": 39.64},
    ]
    procurement = replay_rows(procurement_rows, "采购")
    comprehensive = replay_rows(comprehensive_rows, "综合")
    sales_98 = replay_rows(sales_98_rows, "销售")
    sales_82 = replay_rows(sales_82_rows, "销售")
    production = replay_rows(production_rows, "生产")
    rows = procurement["rows"] + production["rows"] + comprehensive["rows"] + sales_98["rows"] + sales_82["rows"]
    ok = all(abs(row["diff"]) <= 0.2 for row in rows)
    return {
        "ok": ok,
        "procurement": procurement,
        "production": production,
        "comprehensive": comprehensive,
        "sales_docx_98_46": sales_98,
        "sales_prior_82_25": sales_82,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="复原易木平台成绩页评分。")
    parser.add_argument("score_html", nargs="?", type=Path, help="成绩评定 HTML，可省略以运行内置回归。")
    parser.add_argument("--mode", choices=["采购", "生产", "销售", "综合"], default="综合")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()

    if args.self_test or not args.score_html:
        result = built_in_regression()
    else:
        parsed = parse_score_page(args.score_html)
        result = {**replay_rows(parsed["rows"], args.mode), "problem_rows": parsed["problem_rows"], "source": str(args.score_html)}
        result["ok"] = all(abs(row["diff"]) <= 0.2 for row in result["rows"])
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if not result.get("ok", False):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
