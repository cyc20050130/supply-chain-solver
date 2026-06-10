from __future__ import annotations

import argparse
import math
import shutil
import sys
from pathlib import Path
from typing import Any

import win32com.client
import xlrd

import solve
from excel_builtin_solver_model import cell_addr, ensure_solver_xlam_for_com, read_sections_xlrd, solver_run


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


DEFAULT_XLS = Path(r"C:\Users\cyc20\Desktop\excel\电视综合运营★★★☆-标准版个人练习（15_18）场.xls")
DAYS = 30
SHORTAGE_PENALTY = 100_000_000


def log(message: str) -> None:
    print(f"[excel-linked-solver] {message}")


def excel_col(col: int) -> str:
    value = col
    letters = ""
    while value:
        value, rem = divmod(value - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def quoted_sheet(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def ref(sheet_name: str, row: int, col: int) -> str:
    return f"{quoted_sheet(sheet_name)}!${excel_col(col)}${row}"


def rng(sheet_name: str, row: int, col1: int, col2: int) -> str:
    return f"{quoted_sheet(sheet_name)}!${excel_col(col1)}${row}:${excel_col(col2)}${row}"


def source_sheet_name(xls_path: Path) -> str:
    book = xlrd.open_workbook(str(xls_path))
    for idx in range(book.nsheets):
        sheet = book.sheet_by_index(idx)
        preview = " ".join(
            solve.sv(sheet.cell_value(row, col))
            for row in range(min(sheet.nrows, 12))
            for col in range(min(sheet.ncols, 8))
        )
        if "产品及原料清单" in preview:
            return sheet.name
    raise ValueError("没有找到原始题目工作表")


def row_maps(xls_path: Path) -> dict[str, Any]:
    book = xlrd.open_workbook(str(xls_path))
    sheet = None
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
    if sheet is None:
        raise ValueError("没有找到原始题目工作表")

    maps: dict[str, Any] = {
        "products": {},
        "sales": {},
        "suppliers": {},
        "materials": {},
        "factory": {},
        "routes": {},
    }
    for row_idx in range(sheet.nrows):
        row = [sheet.cell_value(row_idx, col) for col in range(sheet.ncols)]
        excel_row = row_idx + 1
        marker = solve.sv(row[1] if len(row) > 1 else "")
        name = solve.sv(row[2] if len(row) > 2 else "")
        if marker in {"产品", "原料"} and name:
            maps["products"][name] = excel_row
        if marker and marker.endswith("总经销"):
            maps["sales"][marker] = excel_row
        if marker and marker.endswith(("厂", "工厂")) and len(row) > 8:
            cargo = solve.sv(row[2])
            if row_idx + 1 >= 33 and row_idx + 1 <= 38:
                maps["suppliers"][marker] = excel_row
        if marker == "电视工厂" and solve.sv(row[2]) == "液晶电视":
            maps["factory"][marker] = excel_row
        if marker == "电视工厂" and solve.sv(row[2]) != "液晶电视":
            maps["materials"][solve.sv(row[2])] = excel_row
        route = solve.sv(row[1] if len(row) > 1 else "")
        if "-->" in route:
            maps["routes"][route] = excel_row
    return maps


def write_row(sheet: Any, row: int, values: list[Any]) -> None:
    for col, value in enumerate(values, start=1):
        sheet.Cells(row, col).Value = value


def delete_sheet_if_exists(workbook: Any, name: str) -> None:
    for sheet in list(workbook.Worksheets):
        if sheet.Name == name:
            sheet.Delete()
            return


def parse_case(xls_path: Path) -> dict[str, Any]:
    sections = read_sections_xlrd(xls_path)
    rates = solve.parse_rates(sections)
    return {
        "products": solve.parse_products(sections),
        "factories": solve.parse_factories(sections),
        "factory_materials": solve.parse_factory_materials(sections),
        "suppliers": solve.parse_suppliers(sections),
        "routes": solve.parse_routes(sections),
        "rates": rates,
        "sales": solve.parse_sales(sections),
        "carriers": solve.parse_carriers_safe(sections, rates),
    }


def build_lanes(case: dict[str, Any], maps: dict[str, Any]) -> list[dict[str, Any]]:
    factory = case["factories"][0]
    product = factory.product or next((p.name for p in case["products"] if p.kind == "产品"), "")
    sales_nodes = {node.node for node in case["sales"]}
    supplier_by_name = {supplier.name: supplier for supplier in case["suppliers"]}
    lanes: list[dict[str, Any]] = []
    for route in case["routes"]:
        route_row = maps["routes"].get(route.route)
        if not route_row:
            continue
        if route.src == factory.name and route.dst in sales_nodes:
            lanes.append(
                {
                    "type": "成品",
                    "cargo": product,
                    "source": route.src,
                    "destination": route.dst,
                    "route": route,
                    "route_row": route_row,
                    "carrier": solve.route_carrier(route, case["carriers"]),
                    "product_row": maps["products"][product],
                    "supplier_row": None,
                }
            )
        supplier = supplier_by_name.get(route.src)
        if supplier and route.dst == factory.name:
            lanes.append(
                {
                    "type": "原料",
                    "cargo": supplier.material,
                    "source": route.src,
                    "destination": route.dst,
                    "route": route,
                    "route_row": route_row,
                    "carrier": solve.route_carrier(route, case["carriers"]),
                    "product_row": maps["products"][supplier.material],
                    "supplier_row": maps["suppliers"][supplier.name],
                }
            )
    return lanes


def initial_solution(case: dict[str, Any], lanes: list[dict[str, Any]]) -> tuple[int, dict[int, tuple[int, int, float]]]:
    forecasts = [solve.forecast_node(node) for node in case["sales"]]
    factory = case["factories"][0]
    production0 = max(
        0,
        min(
            int(factory.daily * DAYS),
            math.ceil(sum(row["forecast"] for row in forecasts) - sum(row["init"] for row in forecasts) - factory.init),
        ),
    )
    initial_by_lane: dict[int, tuple[int, int, float]] = {}

    for forecast in forecasts:
        amount = max(0, math.ceil(forecast["forecast"] - forecast["init"]))
        if amount <= 0:
            continue
        candidates = [
            (idx, lane)
            for idx, lane in enumerate(lanes)
            if lane["type"] == "成品" and lane["destination"] == forecast["node"]
        ]
        if not candidates:
            continue
        idx, lane = min(
            candidates,
            key=lambda pair: solve.route_cost(pair[1]["route"], amount, solve.charge_ratio(case["products"], pair[1]["cargo"])),
        )
        freight = solve.route_cost(lane["route"], amount, solve.charge_ratio(case["products"], lane["cargo"]))
        initial_by_lane[idx] = (amount, 1, freight)

    suppliers = {supplier.name: supplier for supplier in case["suppliers"]}
    for material in case["factory_materials"]:
        amount = max(0, math.ceil(production0 * solve.material_bom(case["products"], material.material) - material.init))
        if amount <= 0:
            continue
        candidates = [
            (idx, lane)
            for idx, lane in enumerate(lanes)
            if lane["type"] == "原料" and lane["cargo"] == material.material
        ]
        if not candidates:
            continue

        def landed_cost(pair: tuple[int, dict[str, Any]]) -> float:
            lane = pair[1]
            supplier = suppliers[lane["source"]]
            purchase = amount * solve.currency_to_cny(supplier.price, supplier.currency, case["rates"])
            freight = solve.route_cost(lane["route"], amount, solve.charge_ratio(case["products"], lane["cargo"]))
            return purchase + freight

        idx, lane = min(candidates, key=landed_cost)
        freight = solve.route_cost(lane["route"], amount, solve.charge_ratio(case["products"], lane["cargo"]))
        initial_by_lane[idx] = (amount, 1, freight)
    return production0, initial_by_lane


def linked_forecast_formula(sheet_name: str, row: int) -> str:
    monthly = rng(sheet_name, row, 6, 11)
    first3 = rng(sheet_name, row, 6, 8)
    last3 = rng(sheet_name, row, 9, 11)
    trend9 = f"(SUM({last3})-SUM({first3}))/9"
    trend12 = f"(SUM({last3})-SUM({first3}))/12"
    return f"=ROUNDUP(SUMPRODUCT({monthly},{{1,2,3,4,5,6}})/21+IF(ABS({trend9})>150,{trend12},{trend9}),0)"


def write_linked_model(xls_path: Path, run_solver: bool = True) -> dict[str, Any]:
    xls_path = xls_path.resolve()
    ensure_solver_xlam_for_com()
    case = parse_case(xls_path)
    maps = row_maps(xls_path)
    src_sheet = source_sheet_name(xls_path)
    lanes = build_lanes(case, maps)
    factory = case["factories"][0]
    factory_row = maps["factory"][factory.name]
    product = factory.product
    initial_production, initial_by_lane = initial_solution(case, lanes)

    backup = xls_path.with_name(f"{xls_path.stem}_原始备份.xls")
    if not backup.exists():
        shutil.copy2(xls_path, backup)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 1
    workbook = None
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
        model.Range("A3").Value = "总运输费"
        model.Range("A4").Value = "总采购成本"
        model.Range("A5").Value = "缺口罚金"
        model.Range("A6").Value = "候选总分"
        model.Range("A7").Value = "生产量变量"
        model.Range("A8").Value = "联动说明"
        model.Range("B8").Value = f"参数全部引用原始工作表：{src_sheet}。改原表数字后，重新运行规划求解即可。"

        var_start = 2
        var_row = var_start
        write_row(model, 1, ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "变量名", "变量值", "变量类型"])
        model.Cells(var_row, 21).Value = "生产量"
        model.Cells(var_row, 22).Value = initial_production
        model.Cells(var_row, 23).Value = "int"
        production_var_row = var_row
        var_row += 1

        lane_var_rows: list[dict[str, int]] = []
        for idx, lane in enumerate(lanes, start=1):
            rows = {}
            qty0, used0, freight0 = initial_by_lane.get(idx - 1, (0, 0, 0.0))
            for suffix, initial, kind in (("运输量", 0, "int"), ("启用", 0, "bin"), ("运费", 0, "cont")):
                model.Cells(var_row, 21).Value = f"L{idx}_{suffix}"
                if suffix == "运输量":
                    model.Cells(var_row, 22).Value = qty0
                elif suffix == "启用":
                    model.Cells(var_row, 22).Value = used0
                elif suffix == "运费":
                    model.Cells(var_row, 22).Value = freight0
                else:
                    model.Cells(var_row, 22).Value = initial
                model.Cells(var_row, 23).Value = kind
                rows[suffix] = var_row
                var_row += 1
            lane_var_rows.append(rows)

        shortage_rows: dict[str, int] = {}
        for node in case["sales"]:
            model.Cells(var_row, 21).Value = f"销售缺口_{node.node}"
            model.Cells(var_row, 22).Value = 0
            model.Cells(var_row, 23).Value = "int"
            shortage_rows[f"market:{node.node}"] = var_row
            var_row += 1
        for material in case["factory_materials"]:
            model.Cells(var_row, 21).Value = f"原料缺口_{material.material}"
            model.Cells(var_row, 22).Value = 0
            model.Cells(var_row, 23).Value = "int"
            shortage_rows[f"material:{material.material}"] = var_row
            var_row += 1
        last_var_row = var_row - 1

        sales_start = 11
        write_row(model, sales_start, ["销售网点", "预测销量", "期初库存", "库存上限", "销售缺口变量", "来源行"])
        for offset, node in enumerate(case["sales"], start=1):
            row = sales_start + offset
            source_row = maps["sales"][node.node]
            model.Cells(row, 1).Formula = f"={ref(src_sheet, source_row, 2)}"
            model.Cells(row, 2).Formula = linked_forecast_formula(src_sheet, source_row)
            model.Cells(row, 3).Formula = f"={ref(src_sheet, source_row, 5)}"
            model.Cells(row, 4).Formula = f"={ref(src_sheet, source_row, 12)}"
            model.Cells(row, 5).Formula = f"={cell_addr(shortage_rows[f'market:{node.node}'], 22)}"
            model.Cells(row, 6).Value = source_row

        material_start = sales_start + len(case["sales"]) + 3
        write_row(model, material_start, ["原料", "BOM", "期初库存", "库存上限", "原料缺口变量", "来源行"])
        for offset, material in enumerate(case["factory_materials"], start=1):
            row = material_start + offset
            product_row = maps["products"][material.material]
            material_row = maps["materials"][material.material]
            model.Cells(row, 1).Formula = f"={ref(src_sheet, material_row, 3)}"
            model.Cells(row, 2).Formula = f"={ref(src_sheet, product_row, 5)}"
            model.Cells(row, 3).Formula = f"={ref(src_sheet, material_row, 5)}"
            model.Cells(row, 4).Formula = f"={ref(src_sheet, material_row, 7)}"
            model.Cells(row, 5).Formula = f"={cell_addr(shortage_rows[f'material:{material.material}'], 22)}"
            model.Cells(row, 6).Value = material_row

        lane_start = material_start + len(case["factory_materials"]) + 4
        write_row(model, lane_start, ["序号", "类型", "货物", "起点", "终点", "路线", "承运商", "计费重", "运价", "起运量", "起运费", "提前期", "单价", "运输量", "启用", "运费变量", "采购成本", "总成本", "来源行"])
        for idx, lane in enumerate(lanes, start=1):
            row = lane_start + idx
            route_row = lane["route_row"]
            product_row = lane["product_row"]
            supplier_row = lane["supplier_row"]
            rows = lane_var_rows[idx - 1]
            write_row(model, row, [idx, lane["type"], lane["cargo"], lane["source"], lane["destination"]])
            model.Cells(row, 6).Formula = f"={ref(src_sheet, route_row, 2)}"
            model.Cells(row, 7).Value = lane["carrier"]
            model.Cells(row, 8).Formula = f"={ref(src_sheet, product_row, 6)}"
            rate_ref = ref(src_sheet, route_row, 8)
            min_qty_ref = ref(src_sheet, route_row, 9)
            model.Cells(row, 9).Formula = f'=IF(ISNUMBER({rate_ref}),{rate_ref},VALUE(LEFT({rate_ref},FIND("元",{rate_ref})-1)))'
            model.Cells(row, 10).Formula = f'=IF(ISNUMBER({min_qty_ref}),{min_qty_ref},VALUE(LEFT({min_qty_ref},MIN(IFERROR(FIND("台",{min_qty_ref}),999),IFERROR(FIND("件",{min_qty_ref}),999))-1)))'
            model.Cells(row, 11).Formula = f"={ref(src_sheet, route_row, 10)}"
            model.Cells(row, 12).Formula = f"={ref(src_sheet, route_row, 11)}"
            model.Cells(row, 13).Formula = "0" if supplier_row is None else f"={ref(src_sheet, supplier_row, 10)}"
            model.Cells(row, 14).Formula = f"={cell_addr(rows['运输量'], 22)}"
            model.Cells(row, 15).Formula = f"={cell_addr(rows['启用'], 22)}"
            model.Cells(row, 16).Formula = f"={cell_addr(rows['运费'], 22)}"
            model.Cells(row, 17).Formula = f"=N{row}*M{row}"
            model.Cells(row, 18).Formula = f"=P{row}+Q{row}"
            model.Cells(row, 19).Value = route_row
        lane_last = lane_start + len(lanes)

        p_cell = cell_addr(production_var_row, 22)
        market_shortage_refs = [cell_addr(row, 22) for key, row in shortage_rows.items() if key.startswith("market:")]
        material_shortage_refs = [cell_addr(row, 22) for key, row in shortage_rows.items() if key.startswith("material:")]
        forecast_total_formula = f"SUM(B{sales_start + 1}:B{sales_start + len(case['sales'])})"
        model.Range("B2").Formula = "=B3+B4+B5"
        model.Range("B3").Formula = f"=SUM(P{lane_start + 1}:P{lane_last})"
        model.Range("B4").Formula = f"=SUM(Q{lane_start + 1}:Q{lane_last})"
        model.Range("B5").Formula = f"={SHORTAGE_PENALTY}*SUM({','.join(market_shortage_refs + material_shortage_refs)})"
        model.Range("B7").Formula = f"={p_cell}"

        metrics_row = lane_last + 3
        write_row(model, metrics_row, ["指标", "值", "指标值", "得分"])
        metric_rows = [
            ("预测偏差率", "0", 0.05, f"=MAX(0,MIN(10,(1-B{metrics_row + 1}/C{metrics_row + 1})*10))"),
            ("单位物流成本", f"=B3/MAX(1,{forecast_total_formula}-SUM({','.join(market_shortage_refs)}))", 110, f"=MAX(0,MIN(20,(1-(B{metrics_row + 2}-C{metrics_row + 2})/(C{metrics_row + 2}*20%))*20))"),
            ("单位采购成本", f"=B4/MAX(1,B7-MAX({','.join(material_shortage_refs)}))", 1600, f"=MAX(0,MIN(20,(1-(B{metrics_row + 3}-C{metrics_row + 3})/(C{metrics_row + 3}*20%))*20))"),
            ("生产满足率", f"=(B7-MAX({','.join(material_shortage_refs)}))/MAX(1,B7)", 1, f"=MAX(0,MIN(15,B{metrics_row + 4}*15))"),
            ("市场满足率", f"=({forecast_total_formula}-SUM({','.join(market_shortage_refs)}))/{forecast_total_formula}", 1, f"=MAX(0,MIN(35,B{metrics_row + 5}*35))"),
        ]
        for offset, (name, formula, target, score_formula) in enumerate(metric_rows, start=1):
            row = metrics_row + offset
            model.Cells(row, 1).Value = name
            model.Cells(row, 2).Formula = formula if formula.startswith("=") else f"={formula}"
            model.Cells(row, 3).Value = target
            model.Cells(row, 4).Formula = score_formula
        model.Range("B6").Formula = f"=SUM(D{metrics_row + 1}:D{metrics_row + len(metric_rows)})"

        constraints_start = metrics_row + len(metric_rows) + 3
        write_row(model, constraints_start, ["约束", "左边", "关系", "右边", "余量"])
        constraint_rows: list[tuple[int, str]] = []
        cr = constraints_start + 1

        def add_constraint(name: str, lhs_formula: str, relation: str, rhs_formula: str) -> None:
            nonlocal cr
            model.Cells(cr, 1).Value = name
            model.Cells(cr, 2).Formula = lhs_formula
            model.Cells(cr, 3).Value = relation
            model.Cells(cr, 4).Formula = rhs_formula
            model.Cells(cr, 5).Formula = f"=D{cr}-B{cr}" if relation == "<=" else f"=B{cr}-D{cr}"
            constraint_rows.append((cr, relation))
            cr += 1

        big_m = 60000
        for idx in range(1, len(lanes) + 1):
            row = lane_start + idx
            add_constraint(f"L{idx} 运输量<=M启用", f"=N{row}", "<=", f"={big_m}*O{row}")
            add_constraint(f"L{idx} 起运量", f"=N{row}", ">=", f"=J{row}*O{row}")
            add_constraint(f"L{idx} 运费>=变动费", f"=P{row}", ">=", f"=I{row}*H{row}*N{row}")
            add_constraint(f"L{idx} 运费>=起运费", f"=P{row}", ">=", f"=K{row}*O{row}")

        for offset, node in enumerate(case["sales"], start=1):
            row = sales_start + offset
            add_constraint(
                f"{node.node} 市场需求",
                f"=C{row}+SUMIFS(N{lane_start + 1}:N{lane_last},B{lane_start + 1}:B{lane_last},\"成品\",E{lane_start + 1}:E{lane_last},\"{node.node}\")+E{row}",
                ">=",
                f"=B{row}",
            )

        add_constraint(
            "工厂成品供给",
            f"=SUMIFS(N{lane_start + 1}:N{lane_last},B{lane_start + 1}:B{lane_last},\"成品\")",
            "<=",
            f"={ref(src_sheet, factory_row, 5)}+{p_cell}",
        )
        add_constraint("生产量<=产能", f"={p_cell}", "<=", f"={ref(src_sheet, factory_row, 6)}*{DAYS}")

        for offset, material in enumerate(case["factory_materials"], start=1):
            row = material_start + offset
            add_constraint(
                f"{material.material} 原料满足生产",
                f"=C{row}+SUMIFS(N{lane_start + 1}:N{lane_last},B{lane_start + 1}:B{lane_last},\"原料\",C{lane_start + 1}:C{lane_last},\"{material.material}\")+E{row}",
                ">=",
                f"={p_cell}*B{row}",
            )

        for supplier in case["suppliers"]:
            supplier_row = maps["suppliers"][supplier.name]
            add_constraint(
                f"{supplier.name} 可供量",
                f"=SUMIFS(N{lane_start + 1}:N{lane_last},D{lane_start + 1}:D{lane_last},\"{supplier.name}\")",
                "<=",
                f"={ref(src_sheet, supplier_row, 9)}",
            )

        model.Columns("A:W").AutoFit()
        model.Activate()
        excel.CalculateFullRebuild()
        workbook.Save()

        status: Any = "未自动运行"

        solver_run(excel, "SolverReset")
        solver_run(excel, "SolverOk", "$B$2", 2, 0, f"$V${var_start}:$V${last_var_row}", 2, "Simplex LP")
        solver_run(excel, "SolverOptions", 300, 100000, 0.000001, True, False, 1, 1, 1, 0.01, True, 0.000001, True)
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

        if run_solver:
            status = solver_run(excel, "SolverSolve", True)
            solver_run(excel, "SolverFinish", 1)
            excel.CalculateFull()
        else:
            workbook.Save()

        result.Range("A1").Value = "Excel 内置规划求解结果"
        write_row(result, 2, ["Solver状态码", status])
        write_row(result, 3, ["参数来源", f"所有参数公式引用：{src_sheet}"])
        if not run_solver:
            write_row(result, 4, ["说明", "已保存 Solver 目标、变量和约束；可在 数据 -> 规划求解 中直接点“求解”。"])
        write_row(result, 5, ["目标函数", model.Range("B2").Value])
        write_row(result, 6, ["总运输费", model.Range("B3").Value])
        write_row(result, 7, ["总采购成本", model.Range("B4").Value])
        write_row(result, 8, ["缺口罚金", model.Range("B5").Value])
        write_row(result, 9, ["候选总分", model.Range("B6").Value])
        write_row(result, 10, ["生产量", model.Range("B7").Value])
        write_row(result, 12, ["评分项", "实际", "指标", "得分"])
        for offset in range(1, len(metric_rows) + 1):
            source_row = metrics_row + offset
            write_row(result, 12 + offset, [model.Cells(source_row, 1).Value, model.Cells(source_row, 2).Value, model.Cells(source_row, 3).Value, model.Cells(source_row, 4).Value])

        out_row = 20
        write_row(result, out_row, ["类型", "货物", "起点", "终点", "路线", "承运商", "运输量", "启用", "运费", "采购成本"])
        out_row += 1
        for row in range(lane_start + 1, lane_last + 1):
            if float(model.Cells(row, 14).Value or 0) <= 0:
                continue
            write_row(result, out_row, [model.Cells(row, col).Value for col in range(2, 8)] + [model.Cells(row, col).Value for col in range(14, 18)])
            out_row += 1
        result.Columns("A:J").AutoFit()

        workbook.Save()
        return {
            "workbook": str(xls_path),
            "backup": str(backup),
            "source_sheet": src_sheet,
            "solver_status_code": status,
            "objective": model.Range("B2").Value,
            "score": model.Range("B6").Value,
            "production": model.Range("B7").Value,
        }
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()


def main() -> None:
    parser = argparse.ArgumentParser(description="在原始 .xls 上建立参数联动的 Excel 内置规划求解模型")
    parser.add_argument("xls", nargs="?", type=Path, default=DEFAULT_XLS)
    parser.add_argument("--no-solve", action="store_true", help="只写入联动模型和可行初始解，不自动调用 SolverSolve")
    args = parser.parse_args()
    result = write_linked_model(args.xls, run_solver=not args.no_solve)
    print(result)


if __name__ == "__main__":
    main()
