from __future__ import annotations

from dataclasses import dataclass
from itertools import product
from typing import Any

import solve


@dataclass(frozen=True)
class Carrier:
    name: str
    mode: str
    unit_rate_cny: float
    start_fee_cny: float


def parse_carriers(sections: dict[str, list[tuple[Any, ...]]], rates: dict[tuple[str, str], float]) -> list[Carrier]:
    carriers: list[Carrier] = []
    for row in solve.find_section(sections, "承运商信息"):
        name = solve.sv(row[1] if len(row) > 1 else "")
        mode = solve.sv(row[2] if len(row) > 2 else "")
        if not name or name == "物流公司" or not mode:
            continue
        currency = solve.sv(row[7] if len(row) > 7 else "CNY") or "CNY"
        fx = rates.get((currency, "CNY"), 1.0)
        carriers.append(
            Carrier(
                name=name,
                mode=mode,
                unit_rate_cny=solve.nv(row[5] if len(row) > 5 else 0) * fx,
                start_fee_cny=solve.nv(row[6] if len(row) > 6 else 0) * fx,
            )
        )
    return carriers


def route_mode_segments(route_name: str) -> list[str]:
    points = [part.strip() for part in route_name.split("-->") if part.strip()]
    if len(points) < 2:
        return ["汽运"]
    modes = []
    for src, dst in zip(points, points[1:]):
        if "火车站" in src and "火车站" in dst:
            modes.append("铁路")
        elif ("码头" in src or "港" in src) and ("码头" in dst or "港" in dst):
            modes.append("海运")
        else:
            modes.append("汽运")
    return modes


def route_points(route_name: str) -> list[str]:
    return [part.strip() for part in route_name.split("-->") if part.strip()]


def is_international_road_segment(src: str, dst: str) -> bool:
    text = f"{src}{dst}"
    foreign_keywords = ("沙特", "达曼", "石油")
    domestic_keywords = ("防城", "贵阳", "贵州", "贵溪", "江西", "达州", "瓮福", "火车站")
    return any(keyword in text for keyword in foreign_keywords) and not all(keyword in text for keyword in domestic_keywords)


def segment_carrier_candidates(mode: str, src: str, dst: str, carriers: list[Carrier]) -> list[Carrier]:
    candidates = [c for c in carriers if c.mode == mode]
    if mode == "汽运":
        international = is_international_road_segment(src, dst)
        filtered = [
            c for c in candidates
            if ("国际" in c.name) == international
        ]
        if filtered:
            candidates = filtered
        elif not international:
            non_international = [c for c in candidates if "国际" not in c.name]
            if non_international:
                candidates = non_international
    if not candidates:
        candidates = [c for c in carriers if c.mode == "汽运" and "国际" not in c.name]
    return candidates


def infer_route_carrier(route: solve.Route, carriers: list[Carrier]) -> str:
    modes = route_mode_segments(route.route)
    points = route_points(route.route)
    if len(modes) == 1 and modes[0] == "汽运":
        target_unit = route.rate / max(route.distance, 1.0)
        src = points[0] if points else route.src
        dst = points[1] if len(points) > 1 else route.dst
        candidates = segment_carrier_candidates("汽运", src, dst, carriers)
        if not candidates:
            return "汽运承运商"
        best = min(
            candidates,
            key=lambda c: (
                abs(c.unit_rate_cny - target_unit),
                abs(c.start_fee_cny - route.min_freight),
            ),
        )
        return best.name

    choices: list[list[Carrier]] = []
    segments = list(zip(points, points[1:])) if len(points) >= 2 else [(route.src, route.dst)]
    for mode, (src, dst) in zip(modes, segments):
        candidates = segment_carrier_candidates(mode, src, dst, carriers)
        if not candidates:
            return "组合承运商"
        choices.append(candidates)

    best_combo = min(
        product(*choices),
        key=lambda combo: (
            abs(sum(c.start_fee_cny for c in combo) - route.min_freight),
            sum(c.unit_rate_cny for c in combo),
        ),
    )
    return "+".join(carrier.name for carrier in best_combo)


def route_carrier_map(xls_path: Any) -> dict[str, str]:
    sections = solve.read_workbook(xls_path)
    rates = solve.parse_rates(sections)
    carriers = parse_carriers(sections, rates)
    return {route.route: infer_route_carrier(route, carriers) for route in solve.parse_routes(sections)}
