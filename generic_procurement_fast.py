from __future__ import annotations

import argparse
import html
import json
import math
from collections import defaultdict
from pathlib import Path
from typing import Any

import solve
from carrier_infer import infer_route_carrier, parse_carriers


DAYS = 30
NEW_ROUTE_PENALTY_SCORE = 300.0
# 易木运输路线的“运输时间”按整条路线到达日计算；不要额外假设到货后再等一天。
PLATFORM_ARRIVAL_BUFFER_DAYS = 0
SHORTAGE_PENALTY = 10_000_000.0
LANE_USE_PENALTY = 0.0
# 同一路线多趟可以在平台用“承运趟数/频率”填报，不能作为真实目标惩罚。
TRIP_USE_PENALTY = 0.0


def cny(value: float, currency: str, rates: dict[tuple[str, str], float]) -> float:
    return value * rates.get((currency, "CNY"), 1.0)


def esc(value: Any) -> str:
    return html.escape("" if value is None else str(value))


def freight(route: solve.Route, amount: float, ratio: float) -> float:
    return max(route.rate * ratio * amount, route.min_freight if amount > 0 else 0)


def route_unit(route: solve.Route, amount: float, ratio: float) -> float:
    return freight(route, max(amount, 1), ratio) / max(amount, 1)


def procurement_target(case_name: str) -> float | None:
    if "白砂糖" in case_name:
        return 415.0
    return None


def cost_score(actual: float, target: float, points: float) -> float:
    if target <= 0:
        return 0.0
    return max(0.0, min(points, (1 - (actual - target) / (target * 0.2)) * points))


def product_output_units(factories: list[solve.FactoryMaterial], products: list[solve.Product]) -> float:
    by_factory: dict[str, list[float]] = defaultdict(list)
    for factory in factories:
        bom = solve.material_bom(products, factory.material)
        if bom > 0:
            by_factory[factory.factory].append(factory.daily * DAYS / bom)
    return sum(values[0] for values in by_factory.values() if values)


def initial_material_cost(
    factories: list[solve.FactoryMaterial],
    suppliers: list[solve.Supplier],
    rates: dict[tuple[str, str], float],
) -> float:
    first_price_by_material: dict[str, float] = {}
    for supplier in suppliers:
        first_price_by_material.setdefault(
            supplier.material,
            cny(supplier.price, supplier.currency, rates),
        )
    return sum(factory.init * first_price_by_material.get(factory.material, 0.0) for factory in factories)


def solve_procurement(xls_path: Path, out_dir: Path) -> dict[str, Any]:
    sections = solve.read_workbook(xls_path)
    products = solve.parse_products(sections)
    suppliers = solve.parse_suppliers(sections)
    factories = solve.parse_factory_materials(sections)
    routes = solve.parse_routes(sections)
    rates = solve.parse_rates(sections)
    carriers = parse_carriers(sections, rates)

    try:
        import pulp
    except ImportError as exc:
        raise RuntimeError("缺少 PuLP，无法执行整数采购优化") from exc

    lane_options: list[tuple[int, solve.Supplier, solve.FactoryMaterial, solve.Route, float, float]] = []
    max_amount = max(sum(s.available for s in suppliers), 1.0)
    for factory in factories:
        ratio = solve.charge_ratio(products, factory.material)
        for supplier in suppliers:
            if supplier.material != factory.material:
                continue
            for route in routes:
                if route.src == supplier.name and route.dst == factory.factory:
                    lane_options.append((len(lane_options), supplier, factory, route, ratio, cny(supplier.price, supplier.currency, rates)))

    model = pulp.LpProblem("procurement_integer_plan", pulp.LpMinimize)
    x: dict[tuple[int, int], Any] = {}
    y: dict[tuple[int, int], Any] = {}
    fcost: dict[tuple[int, int], Any] = {}
    lane_used: dict[int, Any] = {}
    inv: dict[tuple[str, str, int], Any] = {}
    shortage: dict[tuple[str, str, int], Any] = {}

    for lane_id, supplier, factory, route, ratio, unit_price in lane_options:
        lane_used[lane_id] = pulp.LpVariable(f"lane_{lane_id}", lowBound=0, upBound=1, cat="Binary")
        for day in range(1, DAYS + 1):
            key = (lane_id, day)
            x[key] = pulp.LpVariable(f"x_{lane_id}_{day}", lowBound=0, cat="Integer")
            y[key] = pulp.LpVariable(f"y_{lane_id}_{day}", lowBound=0, upBound=1, cat="Binary")
            fcost[key] = pulp.LpVariable(f"freight_{lane_id}_{day}", lowBound=0)
            model += x[key] <= max_amount * y[key]
            model += x[key] >= route.min_qty * y[key]
            model += y[key] <= lane_used[lane_id]
            model += fcost[key] >= route.rate * ratio * x[key]
            model += fcost[key] >= route.min_freight * y[key]

    for factory in factories:
        for day in range(1, DAYS + 1):
            inv_key = (factory.factory, factory.material, day)
            inv[inv_key] = pulp.LpVariable(f"inv_{len(inv)}", lowBound=0, upBound=factory.limit)
            shortage[inv_key] = pulp.LpVariable(f"short_{len(shortage)}", lowBound=0)
            arrivals = []
            for lane_id, supplier, lane_factory, route, ratio, unit_price in lane_options:
                if lane_factory.factory != factory.factory:
                    continue
                if lane_factory.material != factory.material:
                    continue
                ship_day = day - route.lead - PLATFORM_ARRIVAL_BUFFER_DAYS
                if 1 <= ship_day <= DAYS:
                    arrivals.append(x[(lane_id, ship_day)])
            prev_inv = factory.init if day == 1 else inv[(factory.factory, factory.material, day - 1)]
            model += inv[inv_key] == prev_inv + pulp.lpSum(arrivals) - factory.daily + shortage[inv_key]

    for supplier in suppliers:
        related = [
            (lane_id, s, factory, route, ratio, unit_price)
            for lane_id, s, factory, route, ratio, unit_price in lane_options
            if s.name == supplier.name
        ]
        for day in range(1, DAYS + 1):
            shipped_to_day = [
                x[(lane_id, ship_day)]
                for lane_id, s, factory, route, ratio, unit_price in related
                for ship_day in range(1, day + 1)
            ]
            model += pulp.lpSum(shipped_to_day) <= supplier.init + supplier.daily * day
        model += (
            pulp.lpSum(
                x[(lane_id, ship_day)]
                for lane_id, s, factory, route, ratio, unit_price in related
                for ship_day in range(1, DAYS + 1)
            )
            <= supplier.available
        )

    purchase_terms = []
    freight_terms = []
    route_terms = []
    shortage_terms = []
    trip_terms = []
    for lane_id, supplier, factory, route, ratio, unit_price in lane_options:
        route_terms.append(LANE_USE_PENALTY * lane_used[lane_id])
        for day in range(1, DAYS + 1):
            key = (lane_id, day)
            purchase_terms.append(unit_price * x[key])
            freight_terms.append(fcost[key])
            trip_terms.append(TRIP_USE_PENALTY * y[key])
    for factory in factories:
        for day in range(1, DAYS + 1):
            shortage_terms.append(SHORTAGE_PENALTY * shortage[(factory.factory, factory.material, day)])
    model += pulp.lpSum(purchase_terms + freight_terms + route_terms + trip_terms + shortage_terms)
    status_code = model.solve(pulp.PULP_CBC_CMD(msg=False, gapRel=0.001))
    solver_status = pulp.LpStatus[status_code]

    shipments: list[dict[str, Any]] = []
    failures: list[str] = []
    daily_checks: list[dict[str, Any]] = []
    for lane_id, supplier, factory, route, ratio, unit_price in lane_options:
        for day in range(1, DAYS + 1):
            key = (lane_id, day)
            amount = int(round(pulp.value(x[key]) or 0))
            if amount <= 0:
                continue
            shipments.append(
                {
                    "ship_day": day,
                    "arrival_day": day + route.lead,
                    "available_day": day + route.lead + PLATFORM_ARRIVAL_BUFFER_DAYS,
                    "material": factory.material,
                    "supplier": supplier.name,
                    "factory": factory.factory,
                    "amount": amount,
                    "route": route.route,
                    "route_instance": lane_id,
                    "carrier": infer_route_carrier(route, carriers),
                    "lead": route.lead,
                    "unit_price": unit_price,
                    "purchase_cost": amount * unit_price,
                    "freight_cost": float(pulp.value(fcost[key]) or freight(route, amount, ratio)),
                }
            )

    total_shortage = 0.0
    for factory in factories:
        for day in range(1, DAYS + 1):
            short = float(pulp.value(shortage[(factory.factory, factory.material, day)]) or 0.0)
            total_shortage += short
            if short > 1e-6:
                failures.append(f"{factory.factory}-{factory.material} 第{day}天断料 {short:.2f}")
            before = (
                factory.init
                if day == 1
                else float(pulp.value(inv[(factory.factory, factory.material, day - 1)]) or 0.0)
            )
            daily_checks.append(
                {
                    "day": day,
                    "factory": factory.factory,
                    "material": factory.material,
                    "before_consumption": round(before, 4),
                    "daily_need": factory.daily,
                    "ending": round(float(pulp.value(inv[(factory.factory, factory.material, day)]) or 0.0), 4),
                    "shortage": round(short, 4),
                    "overstock": 0.0,
                }
            )

    total_demand = sum(f.daily * DAYS for f in factories)
    net_need = sum(max(0.0, f.daily * DAYS - f.init) for f in factories)
    total_shipped = sum(s["amount"] for s in shipments)
    total_purchase = sum(s["purchase_cost"] for s in shipments)
    total_freight = sum(s["freight_cost"] for s in shipments)
    initial_cost = initial_material_cost(factories, suppliers, rates)
    route_groups = {(s["material"], s["supplier"], s["factory"], s["route"], s.get("carrier", "")) for s in shipments}
    planned_product_units = product_output_units(factories, products)
    raw_to_product = planned_product_units / max(total_demand, 0.001)
    actual_product_units = max(0.0, (total_demand - total_shortage) * raw_to_product)
    platform_total_cost = total_purchase + total_freight + initial_cost
    unit_procurement_cost = platform_total_cost / max(actual_product_units, 0.001)
    satisfaction = max(0.0, min(1.0, (total_demand - total_shortage) / max(total_demand, 0.001)))
    target = procurement_target(xls_path.name)
    score_rows = []
    if target is not None:
        unit_points = cost_score(unit_procurement_cost, target, 60)
        satisfaction_points = satisfaction * 40
        score_rows = [
            {"item": "单位采购成本", "target": target, "actual": unit_procurement_cost, "points": unit_points, "max": 60},
            {"item": "生产满足率", "target": 1.0, "actual": satisfaction, "points": satisfaction_points, "max": 40},
        ]
    earliest_available_day: dict[tuple[str, str], int] = {}
    for lane_id, supplier, factory, route, ratio, unit_price in lane_options:
        candidate_day = 1 + route.lead + PLATFORM_ARRIVAL_BUFFER_DAYS
        key = (factory.factory, factory.material)
        current = earliest_available_day.get(key)
        earliest_available_day[key] = candidate_day if current is None else min(current, candidate_day)
    unavoidable_shortage = 0.0
    for factory in factories:
        first_available = earliest_available_day.get((factory.factory, factory.material), DAYS + 1)
        unavoidable_shortage += max(0.0, factory.daily * (first_available - 1) - factory.init)
    result = {
        "case": xls_path.name,
        "method": "generic procurement MILP: integer shipment + day-by-day inventory simulation",
        "status": solver_status,
        "checks_ok": solver_status in {"Optimal", "Feasible"} and total_shortage <= unavoidable_shortage + 1e-6,
        "failures": failures,
        "total_demand": total_demand,
        "net_need": net_need,
        "total_shipped": total_shipped,
        "purchase_cost": total_purchase,
        "freight_cost": total_freight,
        "initial_material_cost": initial_cost,
        "total_cost": platform_total_cost,
        "route_penalty_score": len(route_groups) * NEW_ROUTE_PENALTY_SCORE,
        "route_penalty_note": "路线/趟次惩罚已关闭；最终分数只按平台评分公式计算。",
        "planned_product_units": planned_product_units,
        "product_units": actual_product_units,
        "raw_shortage": total_shortage,
        "unavoidable_shortage": unavoidable_shortage,
        "unit_procurement_cost": unit_procurement_cost,
        "production_satisfaction": satisfaction,
        "score_rows": score_rows,
        "score": sum(row["points"] for row in score_rows) if score_rows else None,
        "score_note": "本地按平台单位采购成本口径预测，真实分以平台提交后为准。" if score_rows else "未配置该场次评分指标，暂不显示分数。",
        "shipments": sorted(shipments, key=lambda x: (x["ship_day"], x["arrival_day"], x["factory"], x["material"])),
        "daily_checks": daily_checks,
    }

    out_dir.mkdir(parents=True, exist_ok=True)
    json_path = out_dir / f"{xls_path.stem}_procurement_fast_solution.json"
    json_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    html_path = xls_path.with_name(f"{xls_path.stem}_验证方案.html")
    html_path.write_text(render_html(result), encoding="utf-8")
    result["json_path"] = str(json_path)
    result["html_path"] = str(html_path)
    return result


def table(headers: list[str], rows: list[list[Any]], numeric: set[int]) -> str:
    out = ["<table><thead><tr>"]
    out.extend(f"<th>{esc(h)}</th>" for h in headers)
    out.append("</tr></thead><tbody>")
    for row in rows:
        out.append("<tr>")
        for i, value in enumerate(row):
            cls = "num" if i in numeric else ""
            if isinstance(value, float):
                text = f"{value:,.2f}"
            elif isinstance(value, int):
                text = f"{value:,}"
            else:
                text = esc(value)
            out.append(f'<td class="{cls}">{text}</td>')
        out.append("</tr>")
    out.append("</tbody></table>")
    return "\n".join(out)


def fmt_day(value: Any) -> str:
    try:
        return f"1-{int(value):02d}"
    except (TypeError, ValueError):
        return esc(value)


def carrier_class(carrier: str) -> str:
    if "时代物流" in carrier:
        return "carrier-shidai"
    if "中原物流" in carrier:
        return "carrier-zhongyuan"
    if "南方铁路" in carrier or "中铁" in carrier or "西铁" in carrier or "铁路" in carrier:
        return "carrier-rail"
    if "中远海运" in carrier or "海运" in carrier:
        return "carrier-sea"
    if "国际物流" in carrier:
        return "carrier-international"
    if "时速物流" in carrier:
        return "carrier-speed"
    if "顺风物流" in carrier:
        return "carrier-shunfeng"
    if "中联物流" in carrier:
        return "carrier-zhonglian"
    return "carrier-other"


def expanded_transport_rows(rows: list[list[Any]]) -> list[list[Any]]:
    expanded: list[list[Any]] = []
    seq = 1
    for row in rows:
        route, carrier, amount, start_day, trips, interval = row
        expanded.append([seq, route, carrier, amount, start_day, trips, interval])
        seq += 1
    return expanded


def route_mode_class(route: str) -> str:
    has_rail = "火车站" in route
    has_sea = "码头" in route or "港" in route
    if has_rail and has_sea:
        return "mode-mixed"
    if has_rail:
        return "mode-rail"
    if has_sea:
        return "mode-sea"
    return "mode-road"


def transport_table(rows: list[list[Any]]) -> str:
    headers = ["序号", "路线", "承运商", "运输数量", "起运日期", "承运趟数", "几天一趟"]
    out = ["<table class=\"transport\"><thead><tr>"]
    out.extend(f"<th>{esc(h)}</th>" for h in headers)
    out.append("</tr></thead><tbody>")
    for row in expanded_transport_rows(rows):
        route_idx = 1
        carrier_idx = 2
        start_idx = 4
        mode_class = route_mode_class(str(row[route_idx]))
        out.append("<tr>")
        for i, value in enumerate(row):
            if i == 7:
                continue
            cls = "num" if i in {0, 3, 4, 5, 6} else ""
            if i == carrier_idx:
                cls = f"carrier {mode_class} {carrier_class(str(value))}"
            if isinstance(value, float):
                text = f"{value:,.2f}"
            elif isinstance(value, int):
                text = f"{value:,}"
            else:
                text = esc(value)
            if i == start_idx:
                text = fmt_day(value)
            out.append(f'<td class="{cls}">{text}</td>')
        out.append("</tr>")
    out.append("</tbody></table>")
    return "\n".join(out)


def render_html(result: dict[str, Any]) -> str:
    grouped: dict[tuple[str, str, str, str, str], list[dict[str, Any]]] = defaultdict(list)
    purchase_by_company: dict[tuple[str, str], int] = defaultdict(int)
    for s in result["shipments"]:
        grouped[(s["material"], s["supplier"], s["factory"], s["route"], s.get("carrier", ""))].append(s)
        purchase_by_company[(s["supplier"], s["material"])] += int(s["amount"])
    purchase_rows = [[supplier, material, amount] for (supplier, material), amount in sorted(purchase_by_company.items())]
    compact_rows = []
    for (material, supplier, factory, route, carrier), items in sorted(grouped.items(), key=lambda item: (item[0][2], item[0][0], item[0][1], item[0][3], item[0][4])):
        by_amount: dict[int, list[dict[str, Any]]] = defaultdict(list)
        for item in items:
            by_amount[int(round(item["amount"]))].append(item)
        for single_amount, amount_items in sorted(by_amount.items(), key=lambda pair: (-len(pair[1]), pair[0])):
            days = sorted(s["ship_day"] for s in amount_items)
            arrivals = sorted(s["arrival_day"] for s in amount_items)
            intervals = [b - a for a, b in zip(days, days[1:])]
            interval_days = round(sum(intervals) / len(intervals)) if intervals else 0
            compact_rows.append(
                [
                    route,
                    carrier,
                    single_amount,
                    days[0],
                    len(amount_items),
                    interval_days if interval_days else 1,
                ]
            )
    compact_rows.sort(key=lambda row: (int(row[3]), str(row[0]), str(row[1]), int(row[2])))
    score_rows = [
        [row["item"], f"{row['target']:,.2f}" if row["target"] > 1 else f"{row['target']*100:.2f}%", f"{row['actual']:,.2f}" if row["target"] > 1 else f"{row['actual']*100:.2f}%", f"{row['points']:.2f}", f"{row['max']:.0f}"]
        for row in result["score_rows"]
    ]
    score_block = table(["评分项", "指标", "实际", "得分", "满分"], score_rows, {1, 2, 3, 4}) if score_rows else "<p>未配置该场次评分指标，暂不显示分数。</p>"
    status_block = table(
        ["求解状态", "平台时序校验", "候选分", "单位采购成本", "生产满足率", "断料量"],
        [[result["status"], result["checks_ok"], result["score"] if result["score"] is not None else "未配置", f"{result['unit_procurement_cost']:,.2f}", f"{result['production_satisfaction']*100:.2f}%", f"{result.get('raw_shortage', 0):,.2f}"]],
        {2, 3, 5},
    )
    return f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1"><link rel="icon" href="data:,">
<title>{esc(result['case'])} - 通用采购验证方案</title>
<style>
body{{margin:0;padding:24px;font-family:"Segoe UI","Microsoft YaHei",Arial,sans-serif;background:#f6f7f8;color:#202124}}
main{{max-width:1180px;margin:0 auto}} header,section{{background:#fff;border:1px solid #d9dee3;border-radius:8px;padding:18px 22px;margin-bottom:14px;overflow-x:auto}}
h1{{margin:0 0 8px;font-size:22px}} h2{{margin:0 0 12px;font-size:17px}} .meta{{color:#5f6b76;line-height:1.7}}
.status{{display:inline-block;margin-top:12px;border-radius:999px;padding:6px 10px;background:#ecfdf5;color:#166534;border:1px solid #bbf7d0;font-size:13px}}
table{{width:100%;border-collapse:collapse;font-size:14px}} th{{background:#eef2f6;text-align:left;padding:9px 10px;border-bottom:1px solid #cbd5e1;white-space:nowrap}}
td{{padding:9px 10px;border-bottom:1px solid #e5e7eb;vertical-align:top}} td.num{{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}}
tbody tr:nth-child(even) td{{background:#f7f7f7}} tbody tr:nth-child(odd) td{{background:#fff}}
td.carrier{{font-weight:600;text-align:center;white-space:nowrap;border-left:4px solid rgba(0,0,0,.08)}}
td.carrier-shidai{{background:#e8f7ee!important;color:#166534}} td.carrier-zhongyuan{{background:#fff1d6!important;color:#9a4b00}}
td.carrier-rail{{background:#e7f0ff!important;color:#1d4ed8}} td.carrier-sea{{background:#fff5cc!important;color:#92400e}}
td.carrier-international{{background:#ffe7e7!important;color:#b91c1c}} td.carrier-speed{{background:#efe7ff!important;color:#6d28d9}}
td.carrier-shunfeng{{background:#e7fbff!important;color:#0e7490}} td.carrier-zhonglian{{background:#f1f5f9!important;color:#334155}}
td.carrier-other{{background:#eeeeee!important;color:#333}}
</style></head><body><main>
<header><h1>{esc(result['case'])}</h1>
<div class="meta">先看公司采购量，再按运输计划填报表录入平台；联运路线按整条路线一行填写。</div>
</header>
<section><h2>公司采购量</h2>{table(['公司/供应商','原料','采购量'], purchase_rows, {2})}</section>
<section>{transport_table(compact_rows)}</section>
<section><h2>结果状态与分数</h2>{status_block}{score_block}</section>
</main></body></html>"""


def main() -> None:
    parser = argparse.ArgumentParser(description="Fast generic procurement solver for EasyMoo procurement xls files.")
    parser.add_argument("xls", type=Path)
    parser.add_argument("--out-dir", type=Path, default=Path("verified_outputs/procurement"))
    args = parser.parse_args()
    result = solve_procurement(args.xls, args.out_dir)
    print(json.dumps({k: result[k] for k in ["case", "status", "checks_ok", "total_shipped", "unit_procurement_cost", "html_path", "json_path"]}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
