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
from excel_builtin_solver_model import ensure_solver_xlam_for_com, solver_run


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


ROOT = Path(r"C:\Users\cyc20\Desktop\供应链管理")
TARGET_DIR = Path(r"C:\Users\cyc20\Desktop\excel")
DAYS = 30
SHORTAGE_PENALTY = 100_000_000
BIG_M = 1_000_000

SOLVER_STATUS_TEXT = {
    0: "找到满足条件的解",
    1: "已收敛到当前解",
    2: "无法进一步改进当前解",
    3: "达到迭代次数限制",
    4: "未收敛",
    5: "未找到可行解",
    6: "达到最大求解时间",
    7: "线性条件不满足",
    8: "问题规模过大",
    9: "求解器遇到错误",
    10: "用户中断",
    11: "内存不足",
    13: "求解模型错误",
    14: "找到整数可行解",
}


def log(message: str) -> None:
    print(f"[excel-batch] {message}")


def excel_col(col: int) -> str:
    value = col
    letters = ""
    while value:
        value, rem = divmod(value - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def sheet_quote(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def ref(sheet_name: str, row: int, col: int) -> str:
    return f"{sheet_quote(sheet_name)}!${excel_col(col)}${row}"


def rng(sheet_name: str, row: int, col1: int, col2: int) -> str:
    return f"{sheet_quote(sheet_name)}!${excel_col(col1)}${row}:${excel_col(col2)}${row}"


def cell_addr(row: int, col: int) -> str:
    return f"${excel_col(col)}${row}"


def write_row(sheet: Any, row: int, values: list[Any]) -> None:
    for col, value in enumerate(values, start=1):
        sheet.Cells(row, col).Value = value


def delete_sheet_if_exists(workbook: Any, name: str) -> None:
    for sheet in list(workbook.Worksheets):
        if sheet.Name == name:
            sheet.Delete()
            return


def source_sheet_name(xls_path: Path) -> str:
    book = xlrd.open_workbook(str(xls_path))
    for idx in range(book.nsheets):
        sheet = book.sheet_by_index(idx)
        preview = " ".join(
            solve.sv(sheet.cell_value(row, col))
            for row in range(min(sheet.nrows, 18))
            for col in range(min(sheet.ncols, 10))
        )
        if "产品及原料清单" in preview or "运输路线" in preview:
            return sheet.name
    return book.sheet_by_index(0).name


def read_sections_xlrd(xls_path: Path) -> dict[str, list[tuple[Any, ...]]]:
    tmp_dir = Path(tempfile.mkdtemp(prefix="excel_batch_xls_"))
    try:
        ascii_xls = tmp_dir / "input.xls"
        shutil.copy2(xls_path, ascii_xls)
        book = xlrd.open_workbook(str(ascii_xls))
        sheet = book.sheet_by_index(0)
        for idx in range(book.nsheets):
            candidate = book.sheet_by_index(idx)
            preview = " ".join(
                solve.sv(candidate.cell_value(row, col))
                for row in range(min(candidate.nrows, 18))
                for col in range(min(candidate.ncols, 10))
            )
            if "产品及原料清单" in preview or "运输路线" in preview:
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


def row_maps(xls_path: Path) -> dict[str, Any]:
    book = xlrd.open_workbook(str(xls_path))
    sheet = None
    for idx in range(book.nsheets):
        candidate = book.sheet_by_index(idx)
        preview = " ".join(
            solve.sv(candidate.cell_value(row, col))
            for row in range(min(candidate.nrows, 18))
            for col in range(min(candidate.ncols, 10))
        )
        if "产品及原料清单" in preview or "运输路线" in preview:
            sheet = candidate
            break
    if sheet is None:
        sheet = book.sheet_by_index(0)

    maps: dict[str, Any] = {
        "products": {},
        "sales": {},
        "suppliers": {},
        "materials": {},
        "factories": {},
        "routes": {},
    }
    section = ""
    last_kind = ""
    for row_idx in range(sheet.nrows):
        row = [sheet.cell_value(row_idx, col) for col in range(sheet.ncols)]
        excel_row = row_idx + 1
        row_text = " ".join(solve.sv(cell) for cell in row if solve.sv(cell))
        for marker in maps.keys():
            pass
        if "产品及原料清单" in row_text:
            section = "products"
            last_kind = ""
            continue
        if (
            ("销售网点" in row_text and "销量" in row_text)
            or "销售网点销售数据" in row_text
            or "销售网点历史销售数据" in row_text
            or "销售网点历史数据" in row_text
        ):
            section = "sales"
            continue
        if "供应商产能" in row_text:
            section = "suppliers"
            continue
        if "工厂原料消耗" in row_text:
            section = "materials"
            continue
        if "工厂产能" in row_text:
            section = "factories"
            continue
        if "运输路线" == solve.sv(row[1] if len(row) > 1 else ""):
            section = "routes"
            continue
        if any(title in row_text for title in ("货币汇率表", "承运商信息", "地点信息")):
            section = ""

        col_b = solve.sv(row[1] if len(row) > 1 else "")
        col_c = solve.sv(row[2] if len(row) > 2 else "")
        if section == "products":
            if col_b in {"产品", "原料"}:
                last_kind = col_b
            if col_c and col_c not in {"名称", "货物"} and (col_b in {"产品", "原料"} or last_kind == "原料"):
                maps["products"][col_c] = excel_row
        elif section == "sales":
            if col_b and col_b not in {"销售网点", "货物"} and col_c:
                maps["sales"][col_b] = excel_row
        elif section == "suppliers":
            if col_b and col_b != "供应商" and col_c:
                maps["suppliers"][col_b] = excel_row
        elif section == "materials":
            if col_b and col_b != "工厂" and col_c:
                maps["materials"][(col_b, col_c)] = excel_row
                maps["materials"].setdefault(col_c, excel_row)
        elif section == "factories":
            if col_b and col_b != "工厂" and col_c:
                maps["factories"][(col_b, col_c)] = excel_row
                maps["factories"].setdefault(col_b, excel_row)
        elif section == "routes":
            route = col_b
            if "-->" in route:
                maps["routes"][route] = excel_row
    return maps


def forecast_formula(sheet_name: str, row: int) -> str:
    monthly = rng(sheet_name, row, 6, 11)
    first3 = rng(sheet_name, row, 6, 8)
    last3 = rng(sheet_name, row, 9, 11)
    trend9 = f"(SUM({last3})-SUM({first3}))/9"
    trend12 = f"(SUM({last3})-SUM({first3}))/12"
    return f"=ROUNDUP(SUMPRODUCT({monthly},{{1,2,3,4,5,6}})/21+IF(ABS({trend9})>150,{trend12},{trend9}),0)"


def sales_demand_formula(sheet_name: str, row: int, node: solve.SalesNode) -> str:
    if node.daily:
        return f"=ROUNDUP(SUM({rng(sheet_name, row, 6, 5 + DAYS)}),0)"
    return forecast_formula(sheet_name, row)


def sales_limit_formula(sheet_name: str, row: int, node: solve.SalesNode) -> str:
    if node.daily:
        return f"={ref(sheet_name, row, 5 + DAYS + 1)}"
    return f"={ref(sheet_name, row, 12)}"


def product_for_case(products: list[solve.Product], factories: list[solve.Factory], sales: list[solve.SalesNode]) -> str:
    if factories and factories[0].product:
        return factories[0].product
    if sales and sales[0].product:
        return sales[0].product
    match = next((product for product in products if product.kind == "产品"), None)
    return match.name if match else ""


def build_case(xls_path: Path) -> dict[str, Any]:
    sections = read_sections_xlrd(xls_path)
    rates = solve.parse_rates(sections)
    qtype = solve.detect_type(sections, xls_path)
    return {
        "sections": sections,
        "qtype": qtype,
        "products": solve.parse_products(sections),
        "factories": solve.parse_factories(sections),
        "factory_materials": solve.parse_factory_materials(sections),
        "suppliers": solve.parse_suppliers(sections),
        "routes": solve.parse_routes(sections),
        "rates": rates,
        "sales": solve.parse_sales(sections),
        "carriers": solve.parse_carriers_safe(sections, rates),
        "maps": row_maps(xls_path),
        "source_sheet": source_sheet_name(xls_path),
    }


def route_rate_formula(sheet_name: str, route_row: int) -> str:
    rate_ref = ref(sheet_name, route_row, 8)
    parsed = f'VALUE(LEFT({rate_ref},FIND("元",{rate_ref})-1))'
    return f"=IF(ISNUMBER({rate_ref}),{rate_ref},IF(ISERROR({parsed}),0,{parsed}))"


def route_min_qty_formula(sheet_name: str, route_row: int) -> str:
    min_ref = ref(sheet_name, route_row, 9)
    unit_checks = []
    for unit in ("台", "件", "吨", "千克", "公斤", "箱", "个"):
        found = f'FIND("{unit}",{min_ref})'
        unit_checks.append(f"IF(ISERROR({found}),999,{found})")
    unit_pos = f"MIN({','.join(unit_checks)})"
    parsed = f"VALUE(LEFT({min_ref},{unit_pos}-1))"
    return f"=IF(ISNUMBER({min_ref}),{min_ref},IF(ISERROR({parsed}),0,{parsed}))"


def excel_text(value: str) -> str:
    return '"' + str(value).replace('"', '""') + '"'


def sumproduct_by_criteria(sum_col: str, start_row: int, end_row: int, criteria: list[tuple[str, str]]) -> str:
    if end_row < start_row:
        return "0"
    terms = [f"({col}{start_row}:{col}{end_row}={excel_text(value)})" for col, value in criteria]
    terms.append(f"{sum_col}{start_row}:{sum_col}{end_row}")
    return f"SUMPRODUCT({','.join(terms)})"


def route_formula_columns(model: Any, row: int, case: dict[str, Any], lane: dict[str, Any], x_row: int, y_row: int, f_row: int) -> None:
    src_sheet = case["source_sheet"]
    route_row = lane["route_row"]
    product_row = lane.get("product_row")
    supplier_row = lane.get("supplier_row")
    write_row(model, row, [lane["idx"], lane["type"], lane["cargo"], lane["source"], lane["destination"]])
    model.Cells(row, 6).Formula = f"={ref(src_sheet, route_row, 2)}"
    model.Cells(row, 7).Value = lane.get("carrier", "")
    model.Cells(row, 8).Formula = f"={ref(src_sheet, product_row, 6)}" if product_row else "=1"
    model.Cells(row, 9).Value = lane["route"].rate
    model.Cells(row, 10).Value = lane["route"].min_qty
    model.Cells(row, 11).Value = lane["route"].min_freight
    model.Cells(row, 12).Value = lane["route"].lead
    model.Cells(row, 13).Formula = "0" if supplier_row is None else f"={ref(src_sheet, supplier_row, 10)}"
    model.Cells(row, 14).Formula = f"={cell_addr(x_row, 22)}"
    model.Cells(row, 15).Formula = f"={cell_addr(y_row, 22)}"
    model.Cells(row, 16).Formula = f"={cell_addr(f_row, 22)}"
    model.Cells(row, 17).Formula = f"=N{row}*M{row}"
    model.Cells(row, 18).Formula = f"=P{row}+Q{row}"
    model.Cells(row, 19).Value = route_row


def initial_lane_quantities(case: dict[str, Any], lanes: list[dict[str, Any]]) -> dict[int, tuple[int, int, float]]:
    initial: dict[int, tuple[int, int, float]] = {}
    products = case["products"]
    suppliers = {supplier.name: supplier for supplier in case["suppliers"]}
    forecasts = [solve.forecast_node(node) for node in case["sales"]]

    for forecast in forecasts:
        need = max(0, math.ceil(forecast["forecast"] - forecast["init"]))
        candidates = [
            lane for lane in lanes
            if lane["type"] == "成品" and lane["destination"] == forecast["node"]
        ]
        if need > 0 and candidates:
            lane = min(candidates, key=lambda item: solve.route_cost(item["route"], need, solve.charge_ratio(products, item["cargo"])))
            freight = solve.route_cost(lane["route"], need, solve.charge_ratio(products, lane["cargo"]))
            initial[lane["idx"] - 1] = (need, 1, freight)

    production_guess = sum(qty for idx, (qty, _used, _freight) in initial.items() if lanes[idx]["type"] == "成品")
    if not production_guess and forecasts:
        production_guess = sum(row["forecast"] for row in forecasts)
    for material in case["factory_materials"]:
        need = max(0, math.ceil(production_guess * solve.material_bom(products, material.material) - material.init))
        candidates = [
            lane for lane in lanes
            if lane["type"] == "原料" and lane["cargo"] == material.material and lane["destination"] == material.factory
        ]
        if need > 0 and candidates:
            def landed(lane: dict[str, Any]) -> float:
                supplier = suppliers.get(lane["source"])
                purchase = need * solve.currency_to_cny(supplier.price, supplier.currency, case["rates"]) if supplier else 0
                freight = solve.route_cost(lane["route"], need, solve.charge_ratio(products, lane["cargo"]))
                return purchase + freight
            lane = min(candidates, key=landed)
            freight = solve.route_cost(lane["route"], need, solve.charge_ratio(products, lane["cargo"]))
            initial[lane["idx"] - 1] = (need, 1, freight)

    for factory_material in case["factory_materials"]:
        if case["qtype"] == "采购":
            need = max(0, math.ceil(factory_material.daily * DAYS - factory_material.init))
            candidates = [
                lane for lane in lanes
                if lane["type"] == "原料" and lane["cargo"] == factory_material.material and lane["destination"] == factory_material.factory
            ]
            if need > 0 and candidates:
                lane = min(candidates, key=lambda item: solve.route_cost(item["route"], need, solve.charge_ratio(products, item["cargo"])))
                freight = solve.route_cost(lane["route"], need, solve.charge_ratio(products, lane["cargo"]))
                initial[lane["idx"] - 1] = (need, 1, freight)
    return initial


def build_lanes(case: dict[str, Any]) -> list[dict[str, Any]]:
    maps = case["maps"]
    products = case["products"]
    routes = case["routes"]
    factories = case["factories"]
    sales = case["sales"]
    suppliers = case["suppliers"]
    product_name = product_for_case(products, factories, sales)
    supplier_by_name = {supplier.name: supplier for supplier in suppliers}
    factory_names = {factory.name for factory in factories}
    sales_nodes = {node.node for node in sales}
    factory_material_keys = {(item.factory, item.material) for item in case["factory_materials"]}
    lanes: list[dict[str, Any]] = []
    idx = 1

    for route in routes:
        route_row = maps["routes"].get(route.route)
        if not route_row:
            continue
        if route.src in factory_names and route.dst in sales_nodes:
            cargo = next((factory.product for factory in factories if factory.name == route.src), product_name)
            lanes.append({
                "idx": idx,
                "type": "成品",
                "cargo": cargo,
                "source": route.src,
                "destination": route.dst,
                "route": route,
                "route_row": route_row,
                "carrier": solve.route_carrier(route, case["carriers"]),
                "product_row": maps["products"].get(cargo),
                "supplier_row": None,
            })
            idx += 1
        supplier = supplier_by_name.get(route.src)
        if supplier and (route.dst, supplier.material) in factory_material_keys:
            lanes.append({
                "idx": idx,
                "type": "原料",
                "cargo": supplier.material,
                "source": route.src,
                "destination": route.dst,
                "route": route,
                "route_row": route_row,
                "carrier": solve.route_carrier(route, case["carriers"]),
                "product_row": maps["products"].get(supplier.material),
                "supplier_row": maps["suppliers"].get(supplier.name),
            })
            idx += 1
    return lanes


def add_variable(model: Any, var_row: int, name: str, value: float, kind: str) -> int:
    model.Cells(var_row, 21).Value = name
    model.Cells(var_row, 22).Value = value
    model.Cells(var_row, 23).Value = kind
    return var_row + 1


def build_workbook_model(
    xls_path: Path,
    solve_model: bool = False,
    max_time: int = 1800,
    status_override: str | None = None,
) -> dict[str, Any]:
    ensure_solver_xlam_for_com()
    case = build_case(xls_path)
    src_sheet = case["source_sheet"]
    maps = case["maps"]
    qtype = case["qtype"]
    lanes = build_lanes(case)
    initial = initial_lane_quantities(case, lanes)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 1
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(xls_path.resolve()))
        delete_sheet_if_exists(workbook, "求解结果")
        delete_sheet_if_exists(workbook, "Excel规划模型")
        model = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        model.Name = "Excel规划模型"
        result = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        result.Name = "求解结果"

        model.Range("A1").Value = f"{qtype} Excel 内置规划求解模型"
        model.Range("A2").Value = "目标函数"
        model.Range("A3").Value = "总运输费"
        model.Range("A4").Value = "总采购成本"
        model.Range("A5").Value = "缺口罚金"
        model.Range("A6").Value = "候选评分"
        model.Range("A8").Value = "联动说明"
        model.Range("B8").Value = f"参数全部引用原始工作表：{src_sheet}。改原表数字后，重新运行规划求解即可。"

        write_row(model, 1, ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "变量名", "变量值", "变量类型"])
        var_row = 2
        production_var_rows: dict[str, int] = {}
        if qtype in {"生产", "综合"}:
            for factory in case["factories"]:
                frow = maps["factories"].get((factory.name, factory.product)) or maps["factories"].get(factory.name)
                capacity = factory.daily * DAYS
                initial_prod = min(capacity, max(0, sum(q for idx, (q, _u, _f) in initial.items() if lanes[idx]["type"] == "成品" and lanes[idx]["source"] == factory.name) - factory.init))
                production_var_rows[factory.name] = var_row
                var_row = add_variable(model, var_row, f"生产量_{factory.name}", initial_prod, "int")

        lane_var_rows: list[dict[str, int]] = []
        for lane in lanes:
            qty0, used0, freight0 = initial.get(lane["idx"] - 1, (0, 0, 0.0))
            rows: dict[str, int] = {"x": var_row}
            var_row = add_variable(model, var_row, f"L{lane['idx']}_运输量", qty0, "int")
            rows["y"] = var_row
            var_row = add_variable(model, var_row, f"L{lane['idx']}_启用", used0, "bin")
            rows["freight"] = var_row
            var_row = add_variable(model, var_row, f"L{lane['idx']}_运费", freight0, "cont")
            lane_var_rows.append(rows)

        shortage_rows: dict[str, int] = {}
        for node in case["sales"]:
            shortage_rows[f"market:{node.node}"] = var_row
            var_row = add_variable(model, var_row, f"销售缺口_{node.node}", 0, "int")
        for material in case["factory_materials"]:
            shortage_rows[f"material:{material.factory}:{material.material}"] = var_row
            var_row = add_variable(model, var_row, f"原料缺口_{material.factory}_{material.material}", 0, "int")
        last_var_row = var_row - 1

        sales_start = 11
        write_row(model, sales_start, ["销售网点", "预测销量", "期初库存", "库存上限", "缺口变量", "来源行"])
        for offset, node in enumerate(case["sales"], start=1):
            row = sales_start + offset
            source_row = maps["sales"].get(node.node)
            if not source_row:
                continue
            model.Cells(row, 1).Formula = f"={ref(src_sheet, source_row, 2)}"
            model.Cells(row, 2).Formula = sales_demand_formula(src_sheet, source_row, node)
            model.Cells(row, 3).Formula = f"={ref(src_sheet, source_row, 5)}"
            model.Cells(row, 4).Formula = sales_limit_formula(src_sheet, source_row, node)
            model.Cells(row, 5).Formula = f"={cell_addr(shortage_rows[f'market:{node.node}'], 22)}"
            model.Cells(row, 6).Value = source_row

        material_start = sales_start + max(1, len(case["sales"])) + 3
        write_row(model, material_start, ["工厂", "原料", "BOM/日耗", "期初库存", "库存上限", "缺口变量", "来源行"])
        for offset, material in enumerate(case["factory_materials"], start=1):
            row = material_start + offset
            material_row = maps["materials"].get((material.factory, material.material)) or maps["materials"].get(material.material)
            product_row = maps["products"].get(material.material)
            model.Cells(row, 1).Value = material.factory
            model.Cells(row, 2).Formula = f"={ref(src_sheet, material_row, 3)}" if material_row else material.material
            if qtype == "采购":
                model.Cells(row, 3).Formula = f"={ref(src_sheet, material_row, 6)}*{DAYS}" if material_row else material.daily * DAYS
            else:
                model.Cells(row, 3).Formula = f"={ref(src_sheet, product_row, 5)}" if product_row else solve.material_bom(case["products"], material.material)
            model.Cells(row, 4).Formula = f"={ref(src_sheet, material_row, 5)}" if material_row else material.init
            model.Cells(row, 5).Formula = f"={ref(src_sheet, material_row, 7)}" if material_row else material.limit
            model.Cells(row, 6).Formula = f"={cell_addr(shortage_rows[f'material:{material.factory}:{material.material}'], 22)}"
            model.Cells(row, 7).Value = material_row or ""

        factory_start = material_start + max(1, len(case["factory_materials"])) + 3
        write_row(model, factory_start, ["工厂", "产品", "期初库存", "日产能", "30天产能", "生产变量", "来源行"])
        for offset, factory in enumerate(case["factories"], start=1):
            row = factory_start + offset
            frow = maps["factories"].get((factory.name, factory.product)) or maps["factories"].get(factory.name)
            model.Cells(row, 1).Value = factory.name
            model.Cells(row, 2).Formula = f"={ref(src_sheet, frow, 3)}" if frow else factory.product
            model.Cells(row, 3).Formula = f"={ref(src_sheet, frow, 5)}" if frow else factory.init
            model.Cells(row, 4).Formula = f"={ref(src_sheet, frow, 6)}" if frow else factory.daily
            model.Cells(row, 5).Formula = f"=D{row}*{DAYS}"
            if factory.name in production_var_rows:
                model.Cells(row, 6).Formula = f"={cell_addr(production_var_rows[factory.name], 22)}"
            else:
                model.Cells(row, 6).Formula = f"=D{row}*{DAYS}"
            model.Cells(row, 7).Value = frow or ""

        lane_start = factory_start + max(1, len(case["factories"])) + 4
        write_row(model, lane_start, ["序号", "类型", "货物", "起点", "终点", "路线", "承运商", "计费重", "运价", "起运量", "起运费", "提前期", "单价", "运输量", "启用", "运费变量", "采购成本", "总成本", "来源行"])
        for lane, rows in zip(lanes, lane_var_rows):
            row = lane_start + lane["idx"]
            route_formula_columns(model, row, case, lane, rows["x"], rows["y"], rows["freight"])
        lane_last = lane_start + len(lanes)

        market_shortage_refs = [cell_addr(row, 22) for key, row in shortage_rows.items() if key.startswith("market:")]
        material_shortage_refs = [cell_addr(row, 22) for key, row in shortage_rows.items() if key.startswith("material:")]
        shortage_refs = market_shortage_refs + material_shortage_refs
        model.Range("B2").Formula = "=B3+B4+B5"
        model.Range("B3").Formula = f"=SUM(P{lane_start + 1}:P{lane_last})" if lanes else "=0"
        model.Range("B4").Formula = f"=SUM(Q{lane_start + 1}:Q{lane_last})" if lanes else "=0"
        model.Range("B5").Formula = f"={SHORTAGE_PENALTY}*SUM({','.join(shortage_refs)})" if shortage_refs else "=0"

        metrics_row = lane_last + 3
        write_row(model, metrics_row, ["指标", "值", "指标值", "得分"])
        score_last = metrics_row
        material_ship_total = sumproduct_by_criteria("N", lane_start + 1, lane_last, [("B", "原料")])
        if qtype == "采购":
            total_need_formula = f"SUM(C{material_start + 1}:C{material_start + len(case['factory_materials'])})"
            write_row(model, metrics_row + 1, ["单位采购成本"])
            model.Cells(metrics_row + 1, 2).Formula = f"=(B3+B4)/MAX(1,{material_ship_total})" if lanes else "=0"
            model.Cells(metrics_row + 1, 3).Value = 0
            model.Cells(metrics_row + 1, 4).Formula = "=0"
            write_row(model, metrics_row + 2, ["生产满足率"])
            model.Cells(metrics_row + 2, 2).Formula = f"=(MAX(0,{total_need_formula}-SUM({','.join(material_shortage_refs)})))/MAX(1,{total_need_formula})" if material_shortage_refs else "=1"
            model.Cells(metrics_row + 2, 3).Value = 1
            model.Cells(metrics_row + 2, 4).Formula = f"=MAX(0,MIN(40,B{metrics_row + 2}*40))"
            score_last = metrics_row + 2
        elif qtype in {"销售", "生产"}:
            total_forecast = f"SUM(B{sales_start + 1}:B{sales_start + len(case['sales'])})"
            write_row(model, metrics_row + 1, ["单位物流成本"])
            model.Cells(metrics_row + 1, 2).Formula = f"=B3/MAX(1,{total_forecast}-SUM({','.join(market_shortage_refs)}))" if market_shortage_refs else "=0"
            model.Cells(metrics_row + 1, 3).Value = 50 if qtype == "销售" else 190
            model.Cells(metrics_row + 1, 4).Formula = f"=MAX(0,MIN({50 if qtype == '销售' else 60},(1-(B{metrics_row + 1}-C{metrics_row + 1})/(C{metrics_row + 1}*20%))*{50 if qtype == '销售' else 60}))"
            write_row(model, metrics_row + 2, ["市场满足率"])
            model.Cells(metrics_row + 2, 2).Formula = f"=({total_forecast}-SUM({','.join(market_shortage_refs)}))/{total_forecast}" if market_shortage_refs else "=1"
            model.Cells(metrics_row + 2, 3).Value = 1
            model.Cells(metrics_row + 2, 4).Formula = f"=MAX(0,MIN({30 if qtype == '销售' else 40},B{metrics_row + 2}*{30 if qtype == '销售' else 40}))"
            if qtype == "销售":
                write_row(model, metrics_row + 3, ["预测偏差率"])
                model.Cells(metrics_row + 3, 2).Value = 0
                model.Cells(metrics_row + 3, 3).Value = 0.05
                model.Cells(metrics_row + 3, 4).Formula = f"=MAX(0,MIN(20,(1-B{metrics_row + 3}/C{metrics_row + 3})*20))"
                score_last = metrics_row + 3
            else:
                score_last = metrics_row + 2
        else:
            total_forecast = f"SUM(B{sales_start + 1}:B{sales_start + len(case['sales'])})"
            rows = [
                ("预测偏差率", "0", 0.05, f"=MAX(0,MIN(10,(1-B{metrics_row + 1}/C{metrics_row + 1})*10))"),
                ("单位物流成本", f"=B3/MAX(1,{total_forecast}-SUM({','.join(market_shortage_refs)}))", 110, f"=MAX(0,MIN(20,(1-(B{metrics_row + 2}-C{metrics_row + 2})/(C{metrics_row + 2}*20%))*20))"),
                ("单位采购成本", f"=B4/MAX(1,{material_ship_total}-SUM({','.join(material_shortage_refs)}))", 1600, f"=MAX(0,MIN(20,(1-(B{metrics_row + 3}-C{metrics_row + 3})/(C{metrics_row + 3}*20%))*20))"),
                ("生产满足率", f"=({material_ship_total}-SUM({','.join(material_shortage_refs)}))/MAX(1,{material_ship_total})", 1, f"=MAX(0,MIN(15,B{metrics_row + 4}*15))"),
                ("市场满足率", f"=({total_forecast}-SUM({','.join(market_shortage_refs)}))/{total_forecast}", 1, f"=MAX(0,MIN(35,B{metrics_row + 5}*35))"),
            ]
            for offset, (name, formula, target, score_formula) in enumerate(rows, start=1):
                row = metrics_row + offset
                write_row(model, row, [name])
                model.Cells(row, 2).Formula = formula if formula.startswith("=") else f"={formula}"
                model.Cells(row, 3).Value = target
                model.Cells(row, 4).Formula = score_formula
            score_last = metrics_row + len(rows)
        model.Range("B6").Formula = f"=SUM(D{metrics_row + 1}:D{score_last})" if score_last > metrics_row else "=0"

        constraints_start = score_last + 3
        write_row(model, constraints_start, ["约束", "左边", "关系", "右边", "余量"])
        constraint_rows: list[tuple[int, str]] = []
        cr = constraints_start + 1

        def add_constraint(name: str, lhs: str, relation: str, rhs: str) -> None:
            nonlocal cr
            model.Cells(cr, 1).Value = name
            model.Cells(cr, 2).Formula = lhs
            model.Cells(cr, 3).Value = relation
            model.Cells(cr, 4).Formula = rhs
            model.Cells(cr, 5).Formula = f"=D{cr}-B{cr}" if relation == "<=" else f"=B{cr}-D{cr}"
            constraint_rows.append((cr, relation))
            cr += 1

        for lane in lanes:
            row = lane_start + lane["idx"]
            add_constraint(f"L{lane['idx']} 运输量<=M启用", f"=N{row}", "<=", f"={BIG_M}*O{row}")
            add_constraint(f"L{lane['idx']} 起运量", f"=N{row}", ">=", f"=J{row}*O{row}")
            add_constraint(f"L{lane['idx']} 运费>=变动费", f"=P{row}", ">=", f"=I{row}*H{row}*N{row}")
            add_constraint(f"L{lane['idx']} 运费>=起运费", f"=P{row}", ">=", f"=K{row}*O{row}")

        for offset, node in enumerate(case["sales"], start=1):
            row = sales_start + offset
            inbound_finished = sumproduct_by_criteria("N", lane_start + 1, lane_last, [("B", "成品"), ("E", node.node)])
            add_constraint(
                f"{node.node} 市场需求",
                f"=C{row}+{inbound_finished}+E{row}",
                ">=",
                f"=B{row}",
            )

        for offset, factory in enumerate(case["factories"], start=1):
            row = factory_start + offset
            if any(lane["type"] == "成品" and lane["source"] == factory.name for lane in lanes):
                outbound_finished = sumproduct_by_criteria("N", lane_start + 1, lane_last, [("B", "成品"), ("D", factory.name)])
                add_constraint(
                    f"{factory.name} 成品供给",
                    f"={outbound_finished}",
                    "<=",
                    f"=C{row}+F{row}",
                )
            if factory.name in production_var_rows:
                add_constraint(f"{factory.name} 产能", f"=F{row}", "<=", f"=E{row}")

        for offset, material in enumerate(case["factory_materials"], start=1):
            row = material_start + offset
            if qtype == "采购":
                rhs = f"=C{row}"
            else:
                prod_cell = cell_addr(production_var_rows.get(material.factory, 0), 22) if material.factory in production_var_rows else "0"
                rhs = f"={prod_cell}*C{row}"
            inbound_material = sumproduct_by_criteria(
                "N",
                lane_start + 1,
                lane_last,
                [("B", "原料"), ("C", material.material), ("E", material.factory)],
            )
            add_constraint(
                f"{material.factory}-{material.material} 原料需求",
                f"=D{row}+{inbound_material}+F{row}",
                ">=",
                rhs,
            )

        for supplier in case["suppliers"]:
            supplier_row = maps["suppliers"].get(supplier.name)
            if supplier_row:
                supplier_total = sumproduct_by_criteria("N", lane_start + 1, lane_last, [("D", supplier.name)])
                add_constraint(
                    f"{supplier.name} 可供量",
                    f"={supplier_total}",
                    "<=",
                    f"={ref(src_sheet, supplier_row, 9)}",
                )

        model.Columns("A:W").AutoFit()
        model.Activate()
        excel.CalculateFullRebuild()

        solver_run(excel, "SolverReset")
        solver_run(excel, "SolverOk", "$B$2", 2, 0, f"$V$2:$V${last_var_row}", 2, "Simplex LP")
        solver_run(excel, "SolverOptions", max_time, 100000, 0.000001, True, False, 1, 1, 1, 0.01, True, 0.000001, True)
        solver_run(excel, "SolverAdd", f"$V$2:$V${last_var_row}", 3, "0")
        for row in range(2, last_var_row + 1):
            kind = str(model.Cells(row, 23).Value or "")
            if kind == "int":
                solver_run(excel, "SolverAdd", f"$V${row}", 4, "")
            elif kind == "bin":
                solver_run(excel, "SolverAdd", f"$V${row}", 5, "")
        for row, relation in constraint_rows:
            solver_run(excel, "SolverAdd", f"$B${row}", 1 if relation == "<=" else 3, f"$D${row}")

        solver_status: Any = status_override or "未自动运行"
        solver_note = (
            f"{status_override}；已保留 Solver 目标、变量和约束，可在 数据 -> 规划求解 中重新点“求解”。"
            if status_override
            else "已保存 Solver 目标、变量和约束；可在 数据 -> 规划求解 中直接点“求解”。"
        )
        if solve_model and not status_override:
            solver_status = solver_run(excel, "SolverSolve", True)
            try:
                solver_run(excel, "SolverFinish", 1)
            except Exception:
                pass
            excel.CalculateFullRebuild()
            solver_note = f"已自动运行 Solver；状态 {solver_status}: {SOLVER_STATUS_TEXT.get(int(solver_status), '未知状态') if isinstance(solver_status, (int, float)) else solver_status}"

        result.Range("A1").Value = "Excel 内置规划求解结果"
        write_row(result, 2, ["Solver状态码", solver_status])
        write_row(result, 3, ["题型", qtype])
        write_row(result, 4, ["参数来源", f"所有参数公式引用：{src_sheet}"])
        write_row(result, 5, ["说明", solver_note])
        write_row(result, 7, ["目标函数", model.Range("B2").Value])
        write_row(result, 8, ["总运输费", model.Range("B3").Value])
        write_row(result, 9, ["总采购成本", model.Range("B4").Value])
        write_row(result, 10, ["缺口罚金", model.Range("B5").Value])
        write_row(result, 11, ["候选评分", model.Range("B6").Value])
        write_row(result, 13, ["评分项", "实际", "指标", "得分"])
        out = 14
        for row in range(metrics_row + 1, score_last + 1):
            write_row(result, out, [model.Cells(row, col).Value for col in range(1, 5)])
            out += 1
        out += 2
        write_row(result, out, ["类型", "货物", "起点", "终点", "路线", "承运商", "运输量", "启用", "运费", "采购成本"])
        out += 1
        for row in range(lane_start + 1, lane_last + 1):
            if float(model.Cells(row, 14).Value or 0) <= 0:
                continue
            write_row(result, out, [model.Cells(row, col).Value for col in range(2, 8)] + [model.Cells(row, col).Value for col in range(14, 18)])
            out += 1
        result.Columns("A:J").AutoFit()
        workbook.CheckCompatibility = False
        workbook.Save()
        return {
            "file": str(xls_path),
            "qtype": qtype,
            "lanes": len(lanes),
            "variables": last_var_row - 1,
            "constraints": len(constraint_rows),
            "solver_status": solver_status,
            "score": model.Range("B6").Value,
        }
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()


def discover_source_files(root: Path) -> list[Path]:
    return sorted(path for path in root.glob("*/*.xls") if path.parent.name in {"采购", "生产", "销售", "综合"})


def copy_sources(root: Path, target_dir: Path) -> list[Path]:
    target_dir.mkdir(parents=True, exist_ok=True)
    copied: list[Path] = []
    for source in discover_source_files(root):
        target = target_dir / source.name
        shutil.copy2(source, target)
        copied.append(target)
    return copied


def main() -> None:
    parser = argparse.ArgumentParser(description="复制12个供应链xls到桌面excel文件夹，并写入Excel内置规划求解模型")
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--target-dir", type=Path, default=TARGET_DIR)
    parser.add_argument("--source-file", type=Path, default=None, help="只处理单个源xls，复制到target-dir后建模/求解")
    parser.add_argument("--solve", action="store_true", help="写入模型后自动调用 Excel Solver 求解并保存结果")
    parser.add_argument("--max-time", type=int, default=1800, help="每个工作簿 Solver 最大求解秒数")
    parser.add_argument("--status-override", default=None, help="不运行 Solver 时写入求解结果的状态说明")
    args = parser.parse_args()
    ensure_solver_xlam_for_com()
    if args.source_file:
        args.target_dir.mkdir(parents=True, exist_ok=True)
        source = args.source_file.resolve()
        target = args.target_dir / source.name
        shutil.copy2(source, target)
        copied = [target]
    else:
        copied = copy_sources(args.root, args.target_dir)
    results: list[dict[str, Any]] = []
    failures: list[dict[str, str]] = []
    for xls_path in copied:
        try:
            log(f"写入模型: {xls_path.name}")
            results.append(
                build_workbook_model(
                    xls_path,
                    solve_model=args.solve,
                    max_time=args.max_time,
                    status_override=args.status_override,
                )
            )
        except Exception as exc:
            failures.append({"file": str(xls_path), "error": str(exc)})
            log(f"失败: {xls_path.name} | {exc}")
    print("RESULTS")
    for row in results:
        print(row)
    if failures:
        print("FAILURES")
        for row in failures:
            print(row)
        raise SystemExit(1)


if __name__ == "__main__":
    main()
