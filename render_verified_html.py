from __future__ import annotations

import csv
import html
import json
from pathlib import Path
from typing import Any

from carrier_infer import route_carrier_map


ROOT = Path(__file__).resolve().parent
TRANSPORT_HEADERS = ["序号", "路线", "承运商", "运输数量", "起运日期", "承运趟数", "几天一趟"]


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def esc(value: Any) -> str:
    return html.escape("" if value is None else str(value))


def fmt(value: Any) -> str:
    if isinstance(value, float):
        return f"{value:,.2f}"
    if isinstance(value, int):
        return f"{value:,}"
    return esc(value)


def table(headers: list[str], rows: list[list[Any]], numeric: set[int] | None = None) -> str:
    numeric = numeric or set()
    out = ["<table><thead><tr>"]
    out.extend(f"<th>{esc(h)}</th>" for h in headers)
    out.append("</tr></thead><tbody>")
    for row in rows:
        out.append("<tr>")
        for idx, cell in enumerate(row):
            cls = "num" if idx in numeric else ""
            out.append(f'<td class="{cls}">{fmt(cell)}</td>')
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
    if "国顺物流" in carrier:
        return "carrier-guoshun"
    if "兴邦物流" in carrier:
        return "carrier-xingbang"
    if "易达快运" in carrier:
        return "carrier-yida"
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
    out = ["<table class=\"transport\"><thead><tr>"]
    out.extend(f"<th>{esc(h)}</th>" for h in TRANSPORT_HEADERS)
    out.append("</tr></thead><tbody>")
    for row in expanded_transport_rows(rows):
        out.append("<tr>")
        mode_class = route_mode_class(str(row[1]))
        for idx, cell in enumerate(row):
            if idx == 7:
                continue
            cls = "num" if idx in {0, 3, 4, 5, 6} else ""
            if idx == 2:
                cls = f"carrier {mode_class} {carrier_class(str(cell))}"
            text = fmt_day(cell) if idx == 4 else fmt(cell)
            out.append(f'<td class="{cls}">{text}</td>')
        out.append("</tr>")
    out.append("</tbody></table>")
    return "\n".join(out)


def section(title: str, body: str) -> str:
    return f"<section><h2>{esc(title)}</h2>{body}</section>"


def page(title: str, subtitle: str, status: str, score_line: str, sections: list[str]) -> str:
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="icon" href="data:,">
  <title>{esc(title)}</title>
  <style>
    body {{ margin: 0; padding: 24px; font-family: "Segoe UI", "Microsoft YaHei", Arial, sans-serif; background: #f6f7f8; color: #202124; }}
    main {{ max-width: 1180px; margin: 0 auto; }}
    header, section {{ background: #fff; border: 1px solid #d9dee3; border-radius: 8px; padding: 18px 22px; margin-bottom: 14px; overflow-x: auto; }}
    h1 {{ margin: 0 0 8px; font-size: 22px; }}
    h2 {{ margin: 0 0 12px; font-size: 17px; }}
    .meta {{ color: #5f6b76; font-size: 14px; line-height: 1.7; }}
    .status {{ display: inline-block; margin-top: 12px; border-radius: 999px; padding: 6px 10px; background: #ecfdf5; color: #166534; border: 1px solid #bbf7d0; font-size: 13px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
    th {{ background: #eef2f6; text-align: left; padding: 9px 10px; border-bottom: 1px solid #cbd5e1; white-space: nowrap; }}
    td {{ padding: 9px 10px; border-bottom: 1px solid #e5e7eb; vertical-align: top; }}
    td.num {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
    tbody tr:nth-child(even) td {{ background: #f7f7f7; }}
    tbody tr:nth-child(odd) td {{ background: #fff; }}
    td.carrier {{ font-weight: 600; text-align: center; white-space: nowrap; border-left: 4px solid rgba(0,0,0,.08); }}
    td.carrier-shidai {{ background: #e8f7ee !important; color: #166534; }}
    td.carrier-zhongyuan {{ background: #fff1d6 !important; color: #9a4b00; }}
    td.carrier-rail {{ background: #e7f0ff !important; color: #1d4ed8; }}
    td.carrier-sea {{ background: #fff5cc !important; color: #92400e; }}
    td.carrier-international {{ background: #ffe7e7 !important; color: #b91c1c; }}
    td.carrier-speed {{ background: #efe7ff !important; color: #6d28d9; }}
    td.carrier-shunfeng {{ background: #e7fbff !important; color: #0e7490; }}
    td.carrier-zhonglian {{ background: #f1f5f9 !important; color: #334155; }}
    td.carrier-guoshun {{ background: #e0f2fe !important; color: #075985; }}
    td.carrier-xingbang {{ background: #dcfce7 !important; color: #15803d; }}
    td.carrier-yida {{ background: #fae8ff !important; color: #86198f; }}
    td.carrier-other {{ background: #eeeeee !important; color: #333; }}
    tr:last-child td {{ border-bottom: 0; }}
  </style>
</head>
<body>
<main>
  <header>
    <h1>{esc(title)}</h1>
    <div class="meta">{esc(subtitle)}</div>
  </header>
  {"".join(sections)}
  <section><h2>结果状态与分数</h2>{table(["状态", "分数/说明"], [[status, score_line or "未配置"]])}</section>
</main>
</body>
</html>
"""


def render_computer() -> Path:
    data = read_json(ROOT / "verified_outputs/computer/computer-production-15_17-solution.json")
    out = ROOT / "生产/电脑生产★★☆-标准版个人练习（15_17）场_验证方案.html"
    carriers = route_carrier_map(ROOT / "生产/电脑生产★★☆-标准版个人练习（15_17）场.xls")
    route_rows = [
        [
            x["path"],
            carriers.get(x["path"], "组合承运商"),
            x["singleQty"],
            x["startDay"],
            x["trips"],
            x["intervalDays"] if int(x["trips"]) > 1 else 1,
        ]
        for x in data["compactLogisticsPlan"]
    ]
    route_rows.sort(key=lambda row: (int(row[3]), str(row[0]), str(row[1]), int(row[2])))
    html_text = page(
        "电脑生产 15_17 验证方案",
        "只保留平台运输弹窗需要填写的字段；联运路线按整条路线一行填写。",
        f"Optimal；failureCount={len(data['failures'])}",
        "",
        [
            section("运输计划填报表", transport_table(route_rows)),
        ],
    )
    out.write_text(html_text, encoding="utf-8")
    return out


def render_lamp() -> Path:
    data = read_json(ROOT / "verified_outputs/lamp/lamp_sales_15_18_solution.json")
    out = ROOT / "销售/节能灯销售★★☆-标准版个人练习（15_18）场_验证方案.html"
    summary = data["summary"]
    carriers = route_carrier_map(ROOT / "销售/节能灯销售★★☆-标准版个人练习（15_18）场.xls")
    grouped: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for item in data["shipments"]:
        grouped.setdefault((item["outlet"], item["route"]), []).append(item)
    compact_rows = []
    for (outlet, route), items in sorted(grouped.items()):
        by_amount: dict[int, list[dict[str, Any]]] = {}
        for item in items:
            by_amount.setdefault(int(round(item["qty"])), []).append(item)
        for single_amount, amount_items in sorted(by_amount.items(), key=lambda pair: (-len(pair[1]), pair[0])):
            days = sorted(x["ship_day"] for x in amount_items)
            intervals = [b - a for a, b in zip(days, days[1:])]
            interval_text = round(sum(intervals) / len(intervals)) if intervals else 1
            compact_rows.append([route, carriers.get(route, "组合承运商"), single_amount, days[0], len(amount_items), interval_text])
    compact_rows.sort(key=lambda row: (int(row[3]), str(row[0]), str(row[1]), int(row[2])))
    html_text = page(
        "节能灯销售 15_18 验证方案",
        "只保留平台运输弹窗需要填写的字段；联运路线按整条路线一行填写。",
        f"{summary['status']}；checks_ok={not data['checks']['failures']}",
        "",
        [
            section("运输计划填报表", transport_table(compact_rows)),
        ],
    )
    out.write_text(html_text, encoding="utf-8")
    return out


def render_tv() -> Path:
    data = read_json(ROOT / "verified_outputs/tv/tv_15_18_solution/summary.json")
    out = ROOT / "综合/电视综合运营★★★☆-标准版个人练习（15_18）场_验证方案.html"
    val = data["validation"]
    carriers = route_carrier_map(ROOT / "综合/电视综合运营★★★☆-标准版个人练习（15_18）场.xls")
    shipments_csv = ROOT / "verified_outputs/tv/tv_15_18_solution/tv_15_18_shipments.csv"
    grouped: dict[tuple[str, int], list[dict[str, Any]]] = {}
    with shipments_csv.open(newline="", encoding="utf-8") as fh:
        for item in csv.DictReader(fh):
            path = item["path"]
            qty = int(round(float(item["qty"])))
            grouped.setdefault((path, qty), []).append(item)
    route_rows = []
    for (route, single_amount), items in sorted(grouped.items(), key=lambda pair: (pair[0][0], pair[0][1])):
        days = sorted(int(x["ship_day"]) for x in items)
        intervals = [b - a for a, b in zip(days, days[1:])]
        interval_text = round(sum(intervals) / len(intervals)) if intervals else 1
        route_rows.append([route, carriers.get(route, "组合承运商"), single_amount, days[0], len(items), interval_text])
    route_rows.sort(key=lambda row: (int(row[3]), str(row[0]), str(row[1]), int(row[2])))
    html_text = page(
        "电视综合运营 15_18 验证方案",
        "只保留平台运输弹窗需要填写的字段；联运路线按整条路线一行填写。",
        f"{data['status']}；validation.ok={val['ok']}；failures={len(val['failures'])}",
        "",
        [
            section("运输计划填报表", transport_table(route_rows)),
        ],
    )
    out.write_text(html_text, encoding="utf-8")
    return out


def main() -> None:
    for path in [render_computer(), render_lamp(), render_tv()]:
        print(path)


if __name__ == "__main__":
    main()
