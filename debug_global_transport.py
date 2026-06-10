from __future__ import annotations

import sys
from pathlib import Path

import solve


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def run(path: Path, qtype: str) -> None:
    sections = solve.read_workbook(path)
    days = solve.plan_days_for_case(path, qtype)
    products = solve.parse_products(sections)
    factories = solve.parse_factories(sections)
    routes = solve.parse_routes(sections)
    rates = solve.parse_rates(sections)
    carriers = solve.parse_carriers_safe(sections, rates)
    forecasts = [
        solve.forecast_node(node, days, use_bias=solve.forecast_bias_for_case(path, qtype))
        for node in solve.parse_sales(sections, days)
    ]
    cargo = products[0].name if products else factories[0].product
    sources = {
        factory.name: {
            "initial": factory.init,
            "supply": [int(round(factory.daily))] * days,
            "supply_is_capacity": True,
            "limit": factory.limit,
        }
        for factory in factories
    }
    destinations = {
        row["node"]: {
            "initial": row["init"],
            "demand": solve.daily_demands_for_forecast(row, days),
            "limit": row["limit"],
            "excess_fee": row.get("excess_fee", 0.0),
        }
        for row in forecasts
    }
    usable_routes = [route for route in routes if route.src in sources and route.dst in destinations]
    result = solve._solve_day_transport_milp(
        name=f"debug_global_{solve.case_keyword(path)}",
        sources=sources,
        destinations=destinations,
        routes=usable_routes,
        products=products,
        cargo=cargo,
        carriers=carriers,
        days=days,
    )
    shipments = [
        {
            "destination": row["destination"],
            "factory": row["source"],
            "source": row["source"],
            "cargo": row["cargo"],
            "amount": row["amount"],
            "route": row["route"],
            "mode": row["mode"],
            "carrier": row.get("carrier", ""),
            "lead": row["lead"],
            "ship_day": row["ship_day"],
            "arrival_day": row["arrival_day"],
            "freight_cost": row["freight_cost"],
            "note": "全局逐日整数运输优化",
        }
        for row in result["shipments"]
    ]
    shipments = solve.fill_market_shortages_with_fast_routes(
        forecasts=forecasts,
        shipments=shipments,
        assignment={
            row["node"]: min(
                factories,
                key=lambda factory: solve.route_score(
                    solve.pick_best_route(usable_routes, factory.name, row["node"], row["forecast"], solve.charge_ratio(products, cargo)) or usable_routes[0],
                    max(row["forecast"], 1),
                    solve.charge_ratio(products, cargo),
                ),
            )
            for row in forecasts
        },
        routes=routes,
        products=products,
        cargo=cargo,
        carriers=carriers,
        factories=factories,
        days=days,
        max_rounds=2,
    )
    shipments = solve.repair_factory_ship_days(shipments, factories, days)
    replay = solve.daily_market_replay(forecasts, shipments, days)
    freight = sum(float(row.get("freight_cost") or 0) for row in shipments)
    denom = replay["served"] if qtype == "销售" else sum(row["amount"] for row in shipments)
    unit = freight / max(denom, 1)
    print(path.name, "status", result["status"], "rows", len(shipments), "freight", round(freight, 2), "unit", round(unit, 2), "market", round(replay["market_satisfaction"], 4), "short", replay["shortage"])


def main() -> None:
    for path, qtype in [
        (Path("销售") / "速冻水饺销售★★☆-标准版个人练习（10_57）场.xls", "销售"),
        (Path("销售") / "热水器销售★★★-标准版个人练习（10_57）场.xls", "销售"),
        (Path("销售") / "节能灯销售★★☆-标准版个人练习（15_18）场.xls", "销售"),
        (Path("生产") / "电脑生产★★☆-标准版个人练习（15_17）场.xls", "生产"),
        (Path("生产") / "毛衣生产★★★-标准版个人练习（10_57）场.xls", "生产"),
        (Path("生产") / "汉堡生产管理★★☆-标准版个人练习（10_57）场.xls", "生产"),
    ]:
        run(path, qtype)


if __name__ == "__main__":
    main()
