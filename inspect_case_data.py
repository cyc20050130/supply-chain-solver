from __future__ import annotations

import sys
from pathlib import Path

import solve


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


CASES = [
    Path("销售") / "速冻水饺销售★★☆-标准版个人练习（10_57）场.xls",
    Path("销售") / "热水器销售★★★-标准版个人练习（10_57）场.xls",
    Path("销售") / "节能灯销售★★☆-标准版个人练习（15_18）场.xls",
    Path("综合") / "蓄电池供应链运营★★★★-标准版个人练习（10_57）场.xls",
    Path("采购") / "硫磺国际采购★★☆-标准版个人练习（15_16）场.xls",
]


def main() -> None:
    for path in CASES:
        sections = solve.read_workbook(path)
        qtype = solve.detect_type(sections, path)
        days = solve.plan_days_for_case(path, qtype)
        products = solve.parse_products(sections)
        factories = solve.parse_factories(sections)
        sales = solve.parse_sales(sections, days)
        forecasts = [solve.forecast_node(node, days, use_bias=solve.forecast_bias_for_case(path, qtype)) for node in sales]
        suppliers = solve.parse_suppliers(sections)
        materials = solve.parse_factory_materials(sections)
        routes = solve.parse_routes(sections)
        print(f"\n### {path} | qtype={qtype} | days={days}")
        print("products:", [vars(item) for item in products])
        print("factories:", [vars(item) for item in factories])
        print(
            "sales:",
            [
                {
                    "node": row["node"],
                    "forecast": row["forecast"],
                    "init": row["init"],
                    "limit": row["limit"],
                    "daily_avg": round(row["daily_avg"], 2),
                    "method": row["method"],
                }
                for row in forecasts
            ],
        )
        print(
            "totals:",
            {
                "forecast": sum(row["forecast"] for row in forecasts),
                "sales_init": sum(row["init"] for row in forecasts),
                "factory_init": sum(item.init for item in factories),
                "factory_cap": sum(item.daily * days for item in factories),
                "factory_total": sum(item.init + item.daily * days for item in factories),
            },
        )
        print("materials:", [vars(item) for item in materials])
        print("suppliers:", [vars(item) for item in suppliers])
        print("routes:")
        for route in routes:
            print(" ", vars(route))


if __name__ == "__main__":
    main()
