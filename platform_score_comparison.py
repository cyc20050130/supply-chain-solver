from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import solve


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


CASES: list[dict[str, Any]] = [
    {
        "case": "牙膏",
        "mode": "采购",
        "source": "_docx_media/3-牙膏-94分/image8.png",
        "reported_total": 94.78,
        "rows": [
            ("单位采购成本", 2370.0, 2400.63, 56.12),
            ("生产满足率", 1.0, 1.0, 40.0),
        ],
        "ranking_rows": [
            ("单位采购成本", 2349.29, 2400.63, 53.44),
            ("生产满足率", 1.0, 1.0, 40.0),
        ],
    },
    {
        "case": "电脑",
        "mode": "生产",
        "source": "_docx_media/4-电脑-100/image5.png",
        "reported_total": 100.0,
        "rows": [
            ("单位物流成本", 190.0, 163.35, 60.0),
            ("市场满足率", 1.0, 1.0, 40.0),
        ],
    },
    {
        "case": "节能灯",
        "mode": "销售",
        "source": "_docx_media/7-节能灯-98/image4.png",
        "reported_total": 98.46,
        "rows": [
            ("预测偏差率", 0.05, 0.0, 20.0),
            ("单位物流成本", 50.0, 49.15, 50.0),
            ("市场满足率", 1.0, 0.9486, 28.46),
        ],
    },
    {
        "case": "热水器",
        "mode": "销售",
        "source": "_docx_media/9-热水器-100/image1.png",
        "reported_total": 100.0,
        "rows": [
            ("预测偏差率", 0.05, 0.0, 20.0),
            ("单位物流成本", 130.0, 117.13, 50.0),
            ("市场满足率", 1.0, 1.0, 30.0),
        ],
    },
    {
        "case": "速冻水饺",
        "mode": "销售",
        "source": "_docx_media/8-水饺-100/image1.png",
        "reported_total": 100.0,
        "rows": [
            ("预测偏差率", 0.05, 0.0, 20.0),
            ("单位物流成本", 25.0, 23.13, 50.0),
            ("市场满足率", 1.0, 1.0, 30.0),
        ],
    },
    {
        "case": "毛衣",
        "mode": "生产",
        "source": "_docx_media/6-毛衣/image6.png",
        "reported_total": 97.56,
        "rows": [
            ("单位物流成本", 110.0, 106.81, 60.0),
            ("市场满足率", 1.0, 0.9391, 37.56),
        ],
    },
    {
        "case": "汉堡",
        "mode": "生产",
        "source": "_docx_media/5-汉堡/image1.png",
        "reported_total": 94.36,
        "rows": [
            ("单位物流成本", 40.0, 40.75, 54.38),
            ("市场满足率", 1.0, 0.9995, 39.98),
        ],
    },
    {
        "case": "电视",
        "mode": "综合",
        "source": "_docx_media/10-电视机-100/image18.png",
        "reported_total": 100.0,
        "rows": [
            ("预测偏差率", 0.05, 0.0, 10.0),
            ("单位物流成本", 100.0, 90.10, 25.0),
            ("单位采购成本", 1550.0, 1534.61, 25.0),
            ("生产满足率", 1.0, 1.0, 15.0),
            ("市场满足率", 1.0, 1.0, 25.0),
        ],
    },
    {
        "case": "羽绒服",
        "mode": "综合",
        "source": "_docx_media/11-羽绒服-100/image12.png",
        "reported_total": 100.0,
        "status": "待评定页配置复算",
        "rows": [
            ("预测偏差率", 0.05, 0.0, 10.0),
            ("单位物流成本", 25.0, 23.43, 20.0),
            ("单位采购成本", 125.0, 123.53, 20.0),
            ("生产满足率", 1.0, 1.0, 15.0),
            ("市场满足率", 1.0, 1.0, 35.0),
        ],
    },
    {
        "case": "蓄电池",
        "mode": "综合",
        "source": "_docx_media/12-汽车蓄电池/image17.png",
        "reported_total": 92.37,
        "rows": [
            ("预测偏差率", 0.05, 0.0037, 9.26),
            ("单位物流成本", 70.0, 74.12, 14.11),
            ("单位采购成本", 310.0, 311.37, 19.56),
            ("生产满足率", 1.0, 0.9894, 19.79),
            ("市场满足率", 1.0, 0.9884, 29.65),
        ],
    },
    {
        "case": "硫磺",
        "mode": "采购",
        "source": "模拟执行 - 硫磺国际采购 ★★☆ - 标准版 - 易木供应链规划仿真平台.html / 用户成绩截图",
        "reported_total": 23.53,
        "rows": [
            ("单位采购成本", 630.0, 957.90, 0.0),
            ("生产满足率", 1.0, 0.5882, 23.53),
        ],
    },
]


def calc(item: str, target: float, actual: float, max_points: float) -> float:
    if "成本" in item:
        return solve.cost_score(actual, target, max_points)
    if item == "预测偏差率":
        return solve.deviation_score(actual, target, max_points)
    return solve.satisfaction_score(actual, max_points)


def main() -> None:
    output = []
    for case in CASES:
        rows = []
        total = 0.0
        ranking_total = None
        ok = True
        for item, target, actual, reported in case["rows"]:
            max_points = solve.score_points(Path(f"{case['case']}.xls"), case["mode"]).get(item)
            if max_points is None:
                max_points = reported if reported else 0.0
            calculated = calc(item, target, actual, max_points)
            total += calculated
            diff = calculated - reported
            ok = ok and abs(diff) <= 0.2
            rows.append(
                {
                    "item": item,
                    "target": target,
                    "actual": actual,
                    "reported_score": reported,
                    "calculated_score": round(calculated, 4),
                    "diff": round(diff, 4),
                    "max": max_points,
                }
            )
        ranking_rows = []
        if case.get("ranking_rows"):
            ranking_total = 0.0
            for item, target, actual, reported in case["ranking_rows"]:
                max_points = solve.score_points(Path(f"{case['case']}.xls"), case["mode"]).get(item)
                if max_points is None:
                    max_points = reported if reported else 0.0
                calculated = calc(item, target, actual, max_points)
                ranking_total += calculated
                diff = calculated - reported
                ok = ok and abs(diff) <= 0.2
                ranking_rows.append(
                    {
                        "item": item,
                        "target": target,
                        "actual": actual,
                        "reported_score": reported,
                        "calculated_score": round(calculated, 4),
                        "diff": round(diff, 4),
                        "max": max_points,
                    }
                )
        platform_total = (total + ranking_total) / 2 if ranking_total is not None else total
        total_diff = platform_total - case["reported_total"]
        output.append(
            {
                "case": case["case"],
                "mode": case["mode"],
                "source": case["source"],
                "status": case.get("status", "原图成绩页复算"),
                "reported_total": case["reported_total"],
                "calculated_total": round(platform_total, 4),
                "indicator_total": round(total, 4),
                "ranking_total": round(ranking_total, 4) if ranking_total is not None else None,
                "total_diff": round(total_diff, 4),
                "ok": ok and abs(total_diff) <= 0.2,
                "rows": rows,
                "ranking_rows": ranking_rows,
            }
        )
    Path("platform_score_comparison.json").write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"ok": all(row["ok"] for row in output), "cases": output}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
