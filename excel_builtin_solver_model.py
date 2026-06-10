from __future__ import annotations

import argparse
import math
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any

import win32com.client
import xlrd

import solve


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


DAYS = 30
DEFAULT_XLS = Path(r"C:\Users\cyc20\Desktop\excel\电视综合运营★★★☆-标准版个人练习（15_18）场.xls")
SHORTAGE_PENALTY = 100_000_000
OFFICE_SOLVER_XLAM = Path(r"C:\Program Files\Microsoft Office\root\Office16\Library\SOLVER\SOLVER.XLAM")
COM_SOLVER_XLAM = Path.home() / "Documents" / "SOLVER.XLAM"


def log(message: str) -> None:
    print(f"[excel-solver] {message}")


def read_sections_xlrd(xls_path: Path) -> dict[str, list[tuple[Any, ...]]]:
    tmp_dir = Path(tempfile.mkdtemp(prefix="excel_solver_xls_"))
    try:
        ascii_xls = tmp_dir / "input.xls"
        shutil.copy2(xls_path, ascii_xls)
        book = xlrd.open_workbook(str(ascii_xls))
        sheet = book.sheet_by_index(0)
        for idx in range(book.nsheets):
            candidate = book.sheet_by_index(idx)
            preview = " ".join(
                solve.sv(candidate.cell_value(row, col))
                for row in range(min(candidate.nrows, 12))
                for col in range(min(candidate.ncols, 8))
            )
            if "产品及原料清单" in preview:
                sheet = candidate
                break
        raw_rows = [tuple(sheet.cell_value(row, col) for col in range(sheet.ncols)) for row in range(sheet.nrows)]
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    title_keywords = [
        "产品及原料清单",
        "货币汇率表",
        "销售网点历史销售数据",
        "销售网点销售数据",
        "销售网点历史数据",
        "销售网点",
        "供应商产能数据",
        "工厂原料消耗及期初库存",
        "工厂产能及期初库存",
        "工厂产能",
        "承运商信息",
        "地点信息",
        "运输路线",
    ]
    title_idx: dict[str, int] = {}
    for idx, row in enumerate(raw_rows):
        row_text = " ".join(solve.sv(cell) for cell in row if solve.sv(cell))
        for keyword in title_keywords:
            if keyword in row_text and keyword not in title_idx:
                title_idx[keyword] = idx

    sections: dict[str, list[tuple[Any, ...]]] = {}
    ordered = sorted(title_idx.items(), key=lambda item: item[1])
    for pos, (title, start) in enumerate(ordered):
        end = ordered[pos + 1][1] if pos + 1 < len(ordered) else len(raw_rows)
        rows = [raw_rows[i] for i in range(start + 1, end) if any(solve.sv(cell) for cell in raw_rows[i])]
        if rows:
            sections[title] = rows
    return sections


def parse_case(xls_path: Path) -> dict[str, Any]:
    sections = read_sections_xlrd(xls_path)
    products = solve.parse_products(sections)
    factories = solve.parse_factories(sections)
    factory_materials = solve.parse_factory_materials(sections)
    suppliers = solve.parse_suppliers(sections)
    routes = solve.parse_routes(sections)
    rates = solve.parse_rates(sections)
    sales = solve.parse_sales(sections)
    carriers = solve.parse_carriers_safe(sections, rates)
    forecasts = [solve.forecast_node(node) for node in sales]
    if not factories or not forecasts:
        raise ValueError("没有识别到综合案例所需的工厂和销售网点")
    return {
        "sections": sections,
        "products": products,
        "factories": factories,
        "factory_materials": factory_materials,
        "suppliers": suppliers,
        "routes": routes,
        "rates": rates,
        "sales": sales,
        "carriers": carriers,
        "forecasts": forecasts,
    }


def build_lanes(case: dict[str, Any]) -> list[dict[str, Any]]:
    factory = case["factories"][0]
    product = factory.product or next((p.name for p in case["products"] if p.kind == "产品"), "")
    sales_nodes = {row["node"] for row in case["forecasts"]}
    lanes: list[dict[str, Any]] = []
    for route in case["routes"]:
        if route.src == factory.name and route.dst in sales_nodes:
            lanes.append({
                "type": "成品",
                "cargo": product,
                "source": route.src,
                "destination": route.dst,
                "route": route,
                "carrier": solve.route_carrier(route, case["carriers"]),
                "ratio": solve.charge_ratio(case["products"], product),
                "unit_purchase": 0.0,
            })

    supplier_by_name = {supplier.name: supplier for supplier in case["suppliers"]}
    for route in case["routes"]:
        supplier = supplier_by_name.get(route.src)
        if supplier and route.dst == factory.name:
            lanes.append({
                "type": "原料",
                "cargo": supplier.material,
                "source": route.src,
                "destination": route.dst,
                "route": route,
                "carrier": solve.route_carrier(route, case["carriers"]),
                "ratio": solve.charge_ratio(case["products"], supplier.material),
                "unit_purchase": solve.currency_to_cny(supplier.price, supplier.currency, case["rates"]),
            })
    return lanes


def cell_addr(row: int, col: int) -> str:
    letters = ""
    value = col
    while value:
        value, rem = divmod(value - 1, 26)
        letters = chr(65 + rem) + letters
    return f"${letters}${row}"


def delete_sheet_if_exists(workbook: Any, name: str) -> None:
    for sheet in list(workbook.Worksheets):
        if sheet.Name == name:
            sheet.Delete()
            return


def write_row(sheet: Any, row: int, values: list[Any]) -> None:
    for col, value in enumerate(values, start=1):
        sheet.Cells(row, col).Value = value


def solver_run(excel: Any, macro: str, *args: Any) -> Any:
    names = [f"Solver.xlam!{macro}", f"SOLVER.XLAM!{macro}", macro]
    last_error: Exception | None = None
    for name in names:
        try:
            return excel.Run(name, *args)
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"调用 Excel Solver 宏失败: {macro}") from last_error


def ensure_solver_xlam_for_com() -> None:
    """Excel COM resolves SOLVER.XLAM relative to Documents on this machine."""
    if not OFFICE_SOLVER_XLAM.exists():
        raise FileNotFoundError(f"找不到 Office Solver 加载项: {OFFICE_SOLVER_XLAM}")
    if not COM_SOLVER_XLAM.exists():
        shutil.copy2(OFFICE_SOLVER_XLAM, COM_SOLVER_XLAM)


def build_excel_model(xls_path: Path) -> dict[str, Any]:
    xls_path = xls_path.resolve()
    ensure_solver_xlam_for_com()
    case = parse_case(xls_path)
    lanes = build_lanes(case)
    factory = case["factories"][0]
    capacity = int(factory.daily * DAYS)
    forecast_total = sum(row["forecast"] for row in case["forecasts"])
    sales_init = sum(row["init"] for row in case["forecasts"])
    initial_production = max(0, min(capacity, math.ceil(forecast_total - sales_init - factory.init)))

    backup = xls_path.with_name(f"{xls_path.stem}_原始备份.xls")
    if not backup.exists():
        shutil.copy2(xls_path, backup)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        workbook = excel.Workbooks.Open(str(xls_path))
        delete_sheet_if_exists(workbook, "求解结果")
        delete_sheet_if_exists(workbook, "Excel规划模型")

        model = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        model.Name = "Excel规划模型"
        result = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        result.Name = "求解结果"

        model.Range("A1").Value = "综合案例 Excel 内置规划求解模型"
        model.Range("A2").Value = "目标函数"
        model.Range("B2").Value = "待写入"
        model.Range("A3").Value = "总运输费"
        model.Range("A4").Value = "总采购成本"
        model.Range("A5").Value = "缺口罚金"
        model.Range("A6").Value = "候选总分"
        model.Range("A7").Value = "生产量变量"
        model.Range("A8").Value = "说明"
        model.Range("B8").Value = "Excel Solver 最小化：运输费+采购成本+缺口罚金；约束保证库存/供应/产能/起运量。"

        var_start = 2
        var_row = var_start
        variable_rows: dict[str, int] = {}
        write_row(model, 1, ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "变量名", "变量值", "变量类型"])
        model.Cells(var_row, 21).Value = "生产量"
        model.Cells(var_row, 22).Value = initial_production
        model.Cells(var_row, 23).Value = "int"
        variable_rows["P"] = var_row
        var_row += 1

        lane_var_rows: list[dict[str, int]] = []
        for idx, lane in enumerate(lanes, start=1):
            rows = {}
            for suffix, initial, kind in (
                ("运输量", 0, "int"),
                ("启用", 0, "bin"),
                ("运费", 0, "cont"),
            ):
                model.Cells(var_row, 21).Value = f"L{idx}_{suffix}"
                model.Cells(var_row, 22).Value = initial
                model.Cells(var_row, 23).Value = kind
                rows[suffix] = var_row
                var_row += 1
            lane_var_rows.append(rows)

        shortage_rows: dict[str, int] = {}
        for forecast in case["forecasts"]:
            model.Cells(var_row, 21).Value = f"销售缺口_{forecast['node']}"
            model.Cells(var_row, 22).Value = 0
            model.Cells(var_row, 23).Value = "int"
            shortage_rows[f"market:{forecast['node']}"] = var_row
            var_row += 1
        for material in case["factory_materials"]:
            model.Cells(var_row, 21).Value = f"原料缺口_{material.material}"
            model.Cells(var_row, 22).Value = 0
            model.Cells(var_row, 23).Value = "int"
            shortage_rows[f"material:{material.material}"] = var_row
            var_row += 1

        last_var_row = var_row - 1

        sales_start = 11
        write_row(model, sales_start, ["销售网点", "预测销量", "期初库存", "库存上限", "销售缺口变量"])
        for offset, forecast in enumerate(case["forecasts"], start=1):
            node_name = forecast["node"]
            row = sales_start + offset
            model.Cells(row, 1).Value = node_name
            model.Cells(row, 2).Value = forecast["forecast"]
            model.Cells(row, 3).Value = forecast["init"]
            model.Cells(row, 4).Value = forecast["limit"]
            market_shortage_cell = cell_addr(shortage_rows[f"market:{node_name}"], 22)
            model.Cells(row, 5).Formula = f"={market_shortage_cell}"

        material_start = sales_start + len(case["forecasts"]) + 3
        write_row(model, material_start, ["原料", "BOM", "期初库存", "库存上限", "原料缺口变量"])
        for offset, material in enumerate(case["factory_materials"], start=1):
            material_name = material.material
            row = material_start + offset
            model.Cells(row, 1).Value = material_name
            model.Cells(row, 2).Value = solve.material_bom(case["products"], material_name)
            model.Cells(row, 3).Value = material.init
            model.Cells(row, 4).Value = material.limit
            material_shortage_cell = cell_addr(shortage_rows[f"material:{material_name}"], 22)
            model.Cells(row, 5).Formula = f"={material_shortage_cell}"

        lane_start = material_start + len(case["factory_materials"]) + 4
        headers = ["序号", "类型", "货物", "起点", "终点", "路线", "承运商", "计费重", "运价", "起运量", "起运费", "提前期", "单价", "运输量", "启用", "运费变量", "采购成本", "总成本"]
        write_row(model, lane_start, headers)
        for idx, lane in enumerate(lanes, start=1):
            row = lane_start + idx
            route = lane["route"]
            rows = lane_var_rows[idx - 1]
            write_row(model, row, [
                idx,
                lane["type"],
                lane["cargo"],
                lane["source"],
                lane["destination"],
                route.route,
                lane["carrier"],
                lane["ratio"],
                route.rate,
                route.min_qty,
                route.min_freight,
                route.lead,
                lane["unit_purchase"],
            ])
            model.Cells(row, 14).Formula = f"={cell_addr(rows['运输量'], 22)}"
            model.Cells(row, 15).Formula = f"={cell_addr(rows['启用'], 22)}"
            model.Cells(row, 16).Formula = f"={cell_addr(rows['运费'], 22)}"
            model.Cells(row, 17).Formula = f"=N{row}*M{row}"
            model.Cells(row, 18).Formula = f"=P{row}+Q{row}"

        lane_last = lane_start + len(lanes)
        p_cell = cell_addr(variable_rows["P"], 22)
        model.Range("B7").Formula = f"={p_cell}"
        model.Range("B3").Formula = f"=SUM(P{lane_start + 1}:P{lane_last})"
        model.Range("B4").Formula = f"=SUM(Q{lane_start + 1}:Q{lane_last})"
        market_shortage_refs = [cell_addr(row, 22) for key, row in shortage_rows.items() if key.startswith("market:")]
        material_shortage_refs = [cell_addr(row, 22) for key, row in shortage_rows.items() if key.startswith("material:")]
        all_shortage_refs = market_shortage_refs + material_shortage_refs
        model.Range("B5").Formula = f"={SHORTAGE_PENALTY}*SUM({','.join(all_shortage_refs)})"
        model.Range("B2").Formula = "=B3+B4+B5"

        metrics_row = lane_last + 3
        write_row(model, metrics_row, ["指标", "值", "指标值", "得分"])
        metric_names = [
            ("预测偏差率", 0, 0.05, f"=MAX(0,MIN(10,(1-B{metrics_row + 1}/C{metrics_row + 1})*10))"),
            ("单位物流成本", f"=B3/MAX(1,{forecast_total}-SUM({','.join(market_shortage_refs)}))", 110, f"=MAX(0,MIN(20,(1-(B{metrics_row + 2}-C{metrics_row + 2})/(C{metrics_row + 2}*20%))*20))"),
            ("单位采购成本", f"=B4/MAX(1,B7-MAX({','.join(material_shortage_refs)}))", 1600, f"=MAX(0,MIN(20,(1-(B{metrics_row + 3}-C{metrics_row + 3})/(C{metrics_row + 3}*20%))*20))"),
            ("生产满足率", f"=(B7-MAX({','.join(material_shortage_refs)}))/MAX(1,B7)", 1, f"=MAX(0,MIN(15,B{metrics_row + 4}*15))"),
            ("市场满足率", f"=({forecast_total}-SUM({','.join(market_shortage_refs)}))/{forecast_total}", 1, f"=MAX(0,MIN(35,B{metrics_row + 5}*35))"),
        ]
        for offset, (name, value, target, score_formula) in enumerate(metric_names, start=1):
            row = metrics_row + offset
            model.Cells(row, 1).Value = name
            if isinstance(value, str):
                model.Cells(row, 2).Formula = value
            else:
                model.Cells(row, 2).Value = value
            model.Cells(row, 3).Value = target
            model.Cells(row, 4).Formula = score_formula
        model.Range("B6").Formula = f"=SUM(D{metrics_row + 1}:D{metrics_row + len(metric_names)})"

        constraints_start = metrics_row + len(metric_names) + 3
        write_row(model, constraints_start, ["约束", "左边", "关系", "右边", "余量"])
        constraint_rows: list[tuple[int, str]] = []
        cr = constraints_start + 1

        def add_constraint(name: str, lhs_formula: str, relation: str, rhs_formula: str) -> None:
            nonlocal cr
            model.Cells(cr, 1).Value = name
            model.Cells(cr, 2).Formula = lhs_formula
            model.Cells(cr, 3).Value = relation
            model.Cells(cr, 4).Formula = rhs_formula
            if relation == "<=":
                model.Cells(cr, 5).Formula = f"=D{cr}-B{cr}"
            else:
                model.Cells(cr, 5).Formula = f"=B{cr}-D{cr}"
            constraint_rows.append((cr, relation))
            cr += 1

        for idx in range(1, len(lanes) + 1):
            row = lane_start + idx
            add_constraint(f"L{idx} 运输量<=M启用", f"=N{row}", "<=", f"={max(capacity + forecast_total, 1)}*O{row}")
            add_constraint(f"L{idx} 起运量", f"=N{row}", ">=", f"=J{row}*O{row}")
            add_constraint(f"L{idx} 运费>=变动费", f"=P{row}", ">=", f"=I{row}*H{row}*N{row}")
            add_constraint(f"L{idx} 运费>=起运费", f"=P{row}", ">=", f"=K{row}*O{row}")

        for forecast in case["forecasts"]:
            node_name = forecast["node"]
            market_shortage_cell = cell_addr(shortage_rows[f"market:{node_name}"], 22)
            add_constraint(
                f"{node_name} 市场需求",
                f"={forecast['init']}+SUMIFS(N{lane_start + 1}:N{lane_last},B{lane_start + 1}:B{lane_last},\"成品\",E{lane_start + 1}:E{lane_last},\"{node_name}\")+{market_shortage_cell}",
                ">=",
                f"={forecast['forecast']}",
            )
        add_constraint(
            "工厂成品供给",
            f"=SUMIFS(N{lane_start + 1}:N{lane_last},B{lane_start + 1}:B{lane_last},\"成品\")",
            "<=",
            f"={factory.init}+{p_cell}",
        )
        add_constraint("生产量<=产能", f"={p_cell}", "<=", f"={capacity}")

        for material in case["factory_materials"]:
            material_name = material.material
            material_shortage_cell = cell_addr(shortage_rows[f"material:{material_name}"], 22)
            add_constraint(
                f"{material_name} 原料满足生产",
                f"={material.init}+SUMIFS(N{lane_start + 1}:N{lane_last},B{lane_start + 1}:B{lane_last},\"原料\",C{lane_start + 1}:C{lane_last},\"{material_name}\")+{material_shortage_cell}",
                ">=",
                f"={p_cell}*{solve.material_bom(case['products'], material_name)}",
            )

        for supplier in case["suppliers"]:
            add_constraint(
                f"{supplier.name} 可供量",
                f"=SUMIFS(N{lane_start + 1}:N{lane_last},D{lane_start + 1}:D{lane_last},\"{supplier.name}\")",
                "<=",
                f"={supplier.available}",
            )

        model.Columns("A:W").AutoFit()
        model.Activate()

        solver_run(excel, "SolverReset")
        solver_run(excel, "SolverOk", "$B$2", 2, 0, f"$V${var_start}:$V${last_var_row}", 2, "Simplex LP")
        solver_run(excel, "SolverOptions", True, 100, 0.000001, 0.000001, False, False, True, 1, 1, 0.000001, False, False, False)
        solver_run(excel, "SolverAdd", f"$V${var_start}:$V${last_var_row}", 3, "0")
        for row in range(var_start, last_var_row + 1):
            kind = str(model.Cells(row, 23).Value or "")
            if kind == "int":
                solver_run(excel, "SolverAdd", f"$V${row}", 4, "")
            elif kind == "bin":
                solver_run(excel, "SolverAdd", f"$V${row}", 5, "")
        for row, relation in constraint_rows:
            relation_code = 1 if relation == "<=" else 3
            solver_run(excel, "SolverAdd", f"$B${row}", relation_code, f"$D${row}")
        status = solver_run(excel, "SolverSolve", True)
        solver_run(excel, "SolverFinish", 1)

        result.Range("A1").Value = "Excel 内置规划求解结果"
        write_row(result, 2, ["Solver状态码", status])
        write_row(result, 4, ["目标函数", model.Range("B2").Value])
        write_row(result, 5, ["总运输费", model.Range("B3").Value])
        write_row(result, 6, ["总采购成本", model.Range("B4").Value])
        write_row(result, 7, ["缺口罚金", model.Range("B5").Value])
        write_row(result, 8, ["候选总分", model.Range("B6").Value])
        write_row(result, 9, ["生产量", model.Range("B7").Value])

        write_row(result, 11, ["评分项", "实际", "指标", "得分"])
        for offset in range(1, len(metric_names) + 1):
            src_row = metrics_row + offset
            write_row(result, 11 + offset, [model.Cells(src_row, 1).Value, model.Cells(src_row, 2).Value, model.Cells(src_row, 3).Value, model.Cells(src_row, 4).Value])

        out_row = 19
        write_row(result, out_row, ["类型", "货物", "起点", "终点", "路线", "承运商", "运输量", "启用", "运费", "采购成本"])
        out_row += 1
        for row in range(lane_start + 1, lane_last + 1):
            if float(model.Cells(row, 14).Value or 0) <= 0:
                continue
            write_row(result, out_row, [
                model.Cells(row, 2).Value,
                model.Cells(row, 3).Value,
                model.Cells(row, 4).Value,
                model.Cells(row, 5).Value,
                model.Cells(row, 6).Value,
                model.Cells(row, 7).Value,
                model.Cells(row, 14).Value,
                model.Cells(row, 15).Value,
                model.Cells(row, 16).Value,
                model.Cells(row, 17).Value,
            ])
            out_row += 1
        result.Columns("A:J").AutoFit()

        workbook.Save()
        return {
            "workbook": str(xls_path),
            "backup": str(backup),
            "solver_status_code": status,
            "objective": model.Range("B2").Value,
            "score": model.Range("B6").Value,
            "production": model.Range("B7").Value,
        }
    finally:
        try:
            workbook.Close(SaveChanges=True)
        except Exception:
            pass
        excel.Quit()


def main() -> None:
    parser = argparse.ArgumentParser(description="直接在 Excel 文件中建立并运行内置规划求解模型")
    parser.add_argument("xls", nargs="?", type=Path, default=DEFAULT_XLS)
    args = parser.parse_args()
    result = build_excel_model(args.xls)
    print(result)


if __name__ == "__main__":
    main()
