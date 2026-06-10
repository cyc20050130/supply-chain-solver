from __future__ import annotations

import json
import sys
from pathlib import Path

import solve


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def row_summary(result: dict) -> dict:
    return {
        "file": result["xls_path"],
        "case": result.get("case_keyword"),
        "qtype": result.get("qtype"),
        "days": result.get("plan_days"),
        "score": round(float(result.get("score") or 0), 2),
        "unit_logistics": round(float(result.get("unit_logistics") or 0), 2),
        "unit_procurement": round(float(result.get("unit_procurement") or 0), 2),
        "market": round(float(result.get("market_satisfaction") or 0), 4),
        "prod": round(float(result.get("production_satisfaction") or 0), 4),
        "ok": bool((result.get("simulation") or {}).get("ok")),
        "risks": list((result.get("simulation") or {}).get("risks", []))[:12],
        "product_shipments": len(result.get("product_transport") or []),
        "material_shipments": len(result.get("material_transport") or []),
        "score_rows": [
            {
                "item": row.get("item"),
                "target": row.get("target"),
                "actual": row.get("actual"),
                "points": row.get("points"),
                "max": row.get("max"),
            }
            for row in result.get("score_rows", [])
        ],
    }


def main() -> None:
    rows = []
    for xls_path in solve.discover_xls(Path.cwd()):
        result = solve.solve_file(xls_path)
        summary = row_summary(result)
        rows.append(summary)
        print(
            f"[summary] {summary['case']} {summary['qtype']} score={summary['score']} "
            f"unit_logistics={summary['unit_logistics']} unit_proc={summary['unit_procurement']} ok={summary['ok']}",
            flush=True,
        )
    Path("run_summary.json").write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(rows, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
