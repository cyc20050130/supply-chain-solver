from __future__ import annotations

from carrier_infer import Carrier
import os
from pathlib import Path
import score_replay
import solve
import tempfile
import unittest
from pathlib import Path


class PlatformScoreTest(unittest.TestCase):
    def test_production_score_formula_matches_platform_example(self) -> None:
        rows = [
            {"item": "单位物流成本", "target": 190.0, "actual": 163.85, "reported_score": 60.0},
            {"item": "市场满足率", "target": 1.0, "actual": 0.991, "reported_score": 39.64},
        ]

        result = score_replay.replay_rows(rows, "生产")

        self.assertLessEqual(abs(result["total"] - 99.64), 0.05)
        self.assertTrue(all(abs(row["diff"]) <= 0.05 for row in result["rows"]))

    def test_daily_market_simulation_counts_stockout_not_total_supply(self) -> None:
        forecasts = [
            {
                "node": "A门店",
                "forecast": 300,
                "init": 0,
                "limit": 9999,
                "daily_demand": [10] * 30,
            }
        ]
        shipments = [
            {
                "destination": "A门店",
                "amount": 300,
                "ship_day": 1,
                "lead": 15,
                "route": "工厂-->A门店",
            }
        ]

        replay = solve.daily_market_replay(forecasts, shipments)

        self.assertLess(replay["market_satisfaction"], 1.0)
        self.assertEqual(replay["served"], 150)
        self.assertEqual(replay["shortage"], 150)

    def test_tv_comprehensive_uses_case_specific_weights(self) -> None:
        points = solve.score_points(Path("电视综合运营★★★☆-标准版个人练习（15_18）场.xls"), "综合")
        rows = [
            solve.deviation_score(0.0, 0.05, points["预测偏差率"]),
            solve.cost_score(90.10, 100.0, points["单位物流成本"]),
            solve.cost_score(1534.61, 1550.0, points["单位采购成本"]),
            solve.satisfaction_score(1.0, points["生产满足率"]),
            solve.satisfaction_score(1.0, points["市场满足率"]),
        ]

        self.assertEqual(points["单位物流成本"], 25.0)
        self.assertAlmostEqual(sum(rows), 100.0, places=2)

    def test_battery_comprehensive_uses_case_specific_weights(self) -> None:
        points = solve.score_points(Path("蓄电池供应链运营★★★★-标准版个人练习（10_57）场.xls"), "综合")
        rows = [
            solve.deviation_score(0.0037, 0.05, points["预测偏差率"]),
            solve.cost_score(74.12, 70.0, points["单位物流成本"]),
            solve.cost_score(311.37, 310.0, points["单位采购成本"]),
            solve.satisfaction_score(0.9894, points["生产满足率"]),
            solve.satisfaction_score(0.9884, points["市场满足率"]),
        ]

        self.assertEqual(points["生产满足率"], 20.0)
        self.assertLessEqual(abs(sum(rows) - 92.37), 0.05)

    def test_production_table_uses_platform_fill_fields(self) -> None:
        rows = solve.production_plan_rows(
            [{"factory": "东莞工厂", "amount": 4234, "capacity": 4380, "init": 500}],
            30,
        )

        self.assertEqual(rows, [["东莞工厂", "1-01", "1-29", 146]])

    def test_production_table_splits_remainder_into_extra_plan(self) -> None:
        rows = solve.production_plan_rows(
            [{"factory": "测试工厂", "amount": 4235, "capacity": 4380, "init": 0}],
            30,
        )

        self.assertEqual(rows, [["测试工厂", "1-01", "1-29", 146], ["测试工厂", "1-30", "1-30", 1]])

    def test_simulation_uses_declared_production_not_implicit_factory_daily(self) -> None:
        factory = solve.Factory("A工厂", "成品", "件", init=0, daily=100, limit=9999, excess_fee=0)

        simulation = solve.simulate_plan(
            sales_fc=[],
            factories=[factory],
            factory_materials=[],
            production_rows=[{"factory": "A工厂", "product": "成品", "amount": 0, "capacity": 3000, "init": 0}],
            procurement_rows=[],
            product_transport_rows=[{"factory": "A工厂", "source": "A工厂", "destination": "客户", "amount": 50, "ship_day": 1}],
            products=[],
            days=30,
        )

        self.assertFalse(simulation["ok"])
        self.assertTrue(any("成品库存为负" in risk for risk in simulation["risks"]))

    def test_factory_finished_goods_cannot_cover_factory_raw_material_shortage(self) -> None:
        factory = solve.Factory("A工厂", "成品", "件", init=500, daily=100, limit=9999, excess_fee=0)
        material = solve.FactoryMaterial("A工厂", "原料", "千克", init=10, daily=0, limit=9999, excess_fee=0)
        products = [
            solve.Product("产品", "成品", "件", bom=1, charge_ratio=1),
            solve.Product("原料", "原料", "千克", bom=2, charge_ratio=1),
        ]

        simulation = solve.simulate_plan(
            sales_fc=[],
            factories=[factory],
            factory_materials=[material],
            production_rows=[{"factory": "A工厂", "product": "成品", "amount": 20, "capacity": 3000, "init": 500}],
            procurement_rows=[],
            product_transport_rows=[],
            products=products,
            days=30,
        )

        self.assertFalse(simulation["ok"])
        self.assertTrue(any("原料总量不足" in risk for risk in simulation["risks"]))

    def test_sweater_platform_carrier_formula_switches_by_amount(self) -> None:
        carriers = [
            Carrier("顺达汽运", "汽运", 460.0, 0.27, 10000.0),
            Carrier("中通汽运", "汽运", 600.0, 0.32, 7000.0),
            Carrier("中铁快运", "铁路", 300.0, 0.11, 30000.0),
            Carrier("南方速运", "汽运", 390.0, 0.38, 2000.0),
        ]
        route = solve.Route(
            route="苏州工厂-->苏州火车站",
            src="苏州工厂",
            dst="苏州火车站",
            distance=22.0,
            rate=0.0,
            min_qty=0.0,
            min_freight=0.0,
            lead=1,
            currency="CNY",
        )

        cases = [
            (499, "南方速运", 4171.64),
            (938, "中通汽运", 7000.0),
            (1684, "顺达汽运", 10002.96),
        ]
        for amount, expected_carrier, expected_cost in cases:
            with self.subTest(amount=amount):
                option = solve.best_transport_option(route, amount, 1.0, carriers)
                self.assertIsNotNone(option)
                assert option is not None
                self.assertEqual(option.carrier, expected_carrier)
                self.assertAlmostEqual(solve.transport_option_cost(option, amount), expected_cost, places=2)

    def test_partial_multisegment_route_keeps_known_segment_carrier(self) -> None:
        carriers = [
            Carrier("顺达汽运", "汽运", 460.0, 0.27, 10000.0),
            Carrier("中通汽运", "汽运", 600.0, 0.32, 7000.0),
            Carrier("中铁快运", "铁路", 300.0, 0.11, 30000.0),
            Carrier("南方速运", "汽运", 390.0, 0.38, 2000.0),
        ]
        route = solve.Route(
            route="苏州工厂-->苏州火车站-->广州火车站-->毛织品外贸公司",
            src="苏州工厂",
            dst="毛织品外贸公司",
            distance=1198.0,
            rate=137.54,
            min_qty=235.0,
            min_freight=50000.0,
            lead=6,
            currency="CNY",
            segment_distances=(22.0, None, None),
        )

        option = solve.best_transport_option(route, 499, 1.0, carriers)

        self.assertIsNotNone(option)
        assert option is not None
        self.assertTrue(option.carrier.startswith("南方速运+"))
        self.assertEqual(option.lead, 6)

    def test_heatwater_multisegment_residual_does_not_swallow_known_segment_minimum(self) -> None:
        carriers = [
            Carrier("捷达物流", "汽运", 600.0, 0.35, 12000.0),
            Carrier("顺达物流", "汽运", 460.0, 0.45, 8000.0),
            Carrier("中铁快运", "铁路", 300.0, 0.11, 62000.0),
        ]
        route = solve.Route(
            route="苏州工厂-->苏州火车站-->武汉火车站-->武汉总代",
            src="苏州工厂",
            dst="武汉总代",
            distance=671.0,
            rate=105.71,
            min_qty=0.0,
            min_freight=74000.0,
            lead=5,
            currency="CNY",
            segment_distances=(22.0, None, 38.0),
        )

        option = solve.best_transport_option(route, 930, 0.5, carriers)

        self.assertIsNotNone(option)
        assert option is not None
        self.assertEqual(option.carrier, "顺达物流+中铁快运+顺达物流")
        self.assertAlmostEqual(solve.transport_option_cost(option, 930), 74000.0, places=2)

    def test_heatwater_har_distances_override_cross_case_verified_segments(self) -> None:
        previous_context = solve.ACTIVE_HAR_CONTEXT
        carriers = [
            Carrier("捷达物流", "汽运", 400.0, 0.35, 12000.0),
            Carrier("顺达物流", "汽运", 600.0, 0.45, 8000.0),
            Carrier("中铁快运", "铁路", 220.0, 0.13, 50000.0),
            Carrier("南方铁路", "铁路", 280.0, 0.16, 30000.0),
        ]
        route = solve.Route(
            route="东莞工厂-->惠州火车站-->武汉火车站-->武汉总代",
            src="东莞工厂",
            dst="武汉总代",
            distance=913.0,
            rate=137.61,
            min_qty=0.0,
            min_freight=74000.0,
            lead=6,
            currency="CNY",
            segment_distances=(None, None, None),
        )
        solve.ACTIVE_HAR_CONTEXT = {
            "segment_distances": {
                "东莞工厂-->惠州火车站": 48.0,
                "惠州火车站-->武汉火车站": 827.0,
                "武汉火车站-->武汉总代": 38.0,
                "苏州工厂-->苏州火车站": 46.0,
            },
            "segment_days_by_carrier": {},
        }
        try:
            self.assertEqual(solve.segment_distance("苏州工厂-->苏州火车站"), 46.0)
            option = solve.best_transport_option(route, 930, 0.5, carriers)
        finally:
            solve.ACTIVE_HAR_CONTEXT = previous_context

        self.assertIsNotNone(option)
        assert option is not None
        self.assertEqual(option.carrier, "顺达物流+中铁快运+顺达物流")
        self.assertAlmostEqual(solve.transport_option_cost(option, 930), 68044.0, places=2)

    def test_route_option_cache_isolated_by_carrier_numbers(self) -> None:
        previous_context = solve.ACTIVE_HAR_CONTEXT
        solve.ROUTE_TRANSPORT_OPTIONS_CACHE.clear()
        route = solve.Route(
            route="苏州工厂-->北方总代",
            src="苏州工厂",
            dst="北方总代",
            distance=500.0,
            rate=100.0,
            min_qty=0.0,
            min_freight=0.0,
            lead=9,
            currency="CNY",
        )
        carriers_a = [
            Carrier("捷达物流", "汽运", 400.0, 0.35, 12000.0),
            Carrier("顺达物流", "汽运", 600.0, 0.45, 8000.0),
        ]
        carriers_b = [
            Carrier("捷达物流", "汽运", 400.0, 0.80, 12000.0),
            Carrier("顺达物流", "汽运", 600.0, 0.20, 8000.0),
        ]
        try:
            solve.ACTIVE_HAR_CONTEXT = {
                "segment_distances": {"苏州工厂-->北方总代": 500.0},
                "segment_days_by_carrier": {"苏州工厂-->北方总代|||捷达物流": 3},
            }
            option_a = solve.best_transport_option(route, 930, 0.5, carriers_a)
            self.assertIsNotNone(option_a)
            assert option_a is not None
            self.assertEqual(option_a.carrier, "捷达物流")
            self.assertEqual(option_a.lead, 2)

            solve.ACTIVE_HAR_CONTEXT = {
                "segment_distances": {"苏州工厂-->北方总代": 500.0},
                "segment_days_by_carrier": {"苏州工厂-->北方总代|||顺达物流": 7},
            }
            option_b = solve.best_transport_option(route, 930, 0.5, carriers_b)
            self.assertIsNotNone(option_b)
            assert option_b is not None
            self.assertEqual(option_b.carrier, "顺达物流")
            self.assertEqual(option_b.lead, 1)
            self.assertLess(
                solve.transport_option_cost(option_b, 930),
                solve.transport_option_cost(option_a, 930),
            )
        finally:
            solve.ACTIVE_HAR_CONTEXT = previous_context
            solve.ROUTE_TRANSPORT_OPTIONS_CACHE.clear()

    def test_route_option_cache_isolated_by_har_segment_days_when_efficiency_missing(self) -> None:
        previous_context = solve.ACTIVE_HAR_CONTEXT
        solve.ROUTE_TRANSPORT_OPTIONS_CACHE.clear()
        route = solve.Route(
            route="苏州工厂-->北方总代",
            src="苏州工厂",
            dst="北方总代",
            distance=500.0,
            rate=100.0,
            min_qty=0.0,
            min_freight=0.0,
            lead=9,
            currency="CNY",
        )
        carriers = [Carrier("人工确认承运商", "汽运", 0.0, 0.20, 8000.0)]
        try:
            solve.ACTIVE_HAR_CONTEXT = {
                "segment_distances": {"苏州工厂-->北方总代": 500.0},
                "segment_days_by_carrier": {"苏州工厂-->北方总代|||人工确认承运商": 3},
            }
            option_a = solve.best_transport_option(route, 930, 0.5, carriers)
            self.assertIsNotNone(option_a)
            assert option_a is not None
            self.assertEqual(option_a.lead, 3)

            solve.ACTIVE_HAR_CONTEXT = {
                "segment_distances": {"苏州工厂-->北方总代": 500.0},
                "segment_days_by_carrier": {"苏州工厂-->北方总代|||人工确认承运商": 5},
            }
            option_b = solve.best_transport_option(route, 930, 0.5, carriers)
            self.assertIsNotNone(option_b)
            assert option_b is not None
            self.assertEqual(option_b.lead, 5)
        finally:
            solve.ACTIVE_HAR_CONTEXT = previous_context
            solve.ROUTE_TRANSPORT_OPTIONS_CACHE.clear()

    def test_shift_delta_candidates_include_carrier_cost_breakpoints(self) -> None:
        carriers = [
            Carrier("捷达物流", "汽运", 400.0, 0.20, 30000.0),
            Carrier("顺达物流", "汽运", 600.0, 0.45, 8000.0),
        ]
        products = [solve.Product("产品", "热水器", "台", bom=1, charge_ratio=1)]
        route = solve.Route(
            route="苏州工厂-->北方总代",
            src="苏州工厂",
            dst="北方总代",
            distance=500.0,
            rate=100.0,
            min_qty=0.0,
            min_freight=8000.0,
            lead=2,
            currency="CNY",
        )

        deltas = solve.shipment_shift_delta_candidates(
            from_row={"route": route.route, "amount": 1000},
            to_row={"route": route.route, "amount": 0},
            max_delta=700,
            route_by_name={route.route: route},
            products=products,
            cargo="热水器",
            carriers=carriers,
        )

        self.assertIn(300, deltas)
        self.assertIn(500, deltas)

    def test_single_arrival_replacement_switches_to_cheaper_feasible_factory(self) -> None:
        forecasts = [
            {
                "node": "北方总代",
                "forecast": 100,
                "init": 0,
                "limit": 200,
                "daily_demand": [0, 0, 100],
            }
        ]
        factories = [
            solve.Factory("高价工厂", "热水器", "台", init=100, daily=0, limit=9999, excess_fee=0),
            solve.Factory("低价工厂", "热水器", "台", init=100, daily=0, limit=9999, excess_fee=0),
        ]
        products = [solve.Product("产品", "热水器", "台", bom=1, charge_ratio=1)]
        routes = [
            solve.Route("高价工厂-->北方总代", "高价工厂", "北方总代", 10, 100, 0, 0, 2, "CNY"),
            solve.Route("低价工厂-->北方总代", "低价工厂", "北方总代", 10, 10, 0, 0, 2, "CNY"),
        ]
        shipments = [
            {
                "destination": "北方总代",
                "factory": "高价工厂",
                "source": "高价工厂",
                "cargo": "热水器",
                "amount": 100,
                "route": "高价工厂-->北方总代",
                "carrier": "高价承运",
                "lead": 2,
                "ship_day": 1,
                "arrival_day": 3,
                "freight_cost": 10000,
            }
        ]

        polished, status = solve.polish_single_arrival_replacements(
            forecasts=forecasts,
            shipments=shipments,
            factories=factories,
            routes=routes,
            products=products,
            cargo="热水器",
            carriers=None,
            days=3,
        )

        self.assertEqual(status["improvements"], 1)
        self.assertEqual(polished[0]["factory"], "低价工厂")
        self.assertLess(polished[0]["freight_cost"], 10000)

    def test_time_limited_status_is_detected_for_candidate_filtering(self) -> None:
        self.assertTrue(solve.is_time_limited_status({"status": "TimeLimitFeasible"}))
        self.assertTrue(solve.is_time_limited_status({"raw_status": "TimeLimit"}))
        self.assertFalse(solve.is_time_limited_status({"status": "Optimal", "method": "InventoryWindow"}))

    def test_har_context_extracts_segment_distance_and_days(self) -> None:
        har_payload = {
            "log": {
                "entries": [
                    {
                        "request": {"url": "https://www.easymoo.com/ec-scps/view/execute/pro-logistics"},
                        "response": {
                            "status": 200,
                            "content": {
                                "mimeType": "application/json",
                                "text": (
                                    '{"carryPlans":[{"polName":"东莞工厂","podName":"惠州火车站",'
                                    '"podCarrierName":"顺达物流","carryDistance":48,"carryDays":1}]}'
                                ),
                            },
                        },
                    },
                    {
                        "request": {"url": "https://www.easymoo.com/ec-scps/view/execute/pro-logistics"},
                        "response": {
                            "status": 200,
                            "content": {
                                "mimeType": "text/html",
                                "text": (
                                    "DATA_CARRIERS.push({carrierName:'顺达物流',carrierType:'P',"
                                    "efficiency:600,unitPrice:.45,lowestCharge:8000,currencyCode:'CNY'});"
                                ),
                            },
                        },
                    },
                ]
            }
        }
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "sample.har"
            path.write_text(__import__("json").dumps(har_payload, ensure_ascii=False), encoding="utf-8")
            context = solve.load_har_context(path)

        assert context is not None
        self.assertEqual(context["segment_distances"]["东莞工厂-->惠州火车站"], 48.0)
        self.assertEqual(context["segment_days_by_carrier"]["东莞工厂-->惠州火车站|||顺达物流"], 1)
        self.assertEqual(context["carriers"][0]["name"], "顺达物流")

    def test_frontend_html_context_extracts_carriers_and_carry_plans(self) -> None:
        frontend = """
        <script>
        DATA_CARRIERS.push({carrierName:'捷达物流',carrierType:'P',efficiency:400,unitPrice:.35,lowestCharge:12000,currencyCode:'CNY'});
        window.__CASE__ = {"carryPlans":[{"polName":"苏州工厂","podName":"北方总代","podCarrierName":"捷达物流","carryDistance":987,"carryDays":3}]};
        </script>
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "frontend.html"
            path.write_text(frontend, encoding="utf-8")
            context = solve.load_har_context(path)

        assert context is not None
        self.assertEqual(context["source_kind"], "frontend")
        self.assertEqual(context["carriers"][0]["name"], "捷达物流")
        self.assertEqual(context["segment_distances"]["苏州工厂-->北方总代"], 987.0)
        self.assertEqual(context["segment_days_by_carrier"]["苏州工厂-->北方总代|||捷达物流"], 3)

    def test_multiple_frontend_contexts_are_merged_for_one_case(self) -> None:
        transport_frontend = """
        DATA_CARRIERS.push({carrierName:'捷达物流',carrierType:'P',efficiency:400,unitPrice:.35,lowestCharge:12000,currencyCode:'CNY'});
        {"carryPlans":[{"polName":"苏州工厂","podName":"北方总代","podCarrierName":"捷达物流","carryDistance":987,"carryDays":3}]}
        """
        sales_frontend = """
        DATA_CARRIERS.push({carrierName:'捷达物流',carrierType:'P',efficiency:420,unitPrice:.31,lowestCharge:11000,currencyCode:'CNY'});
        {"carryPlans":[{"polName":"苏州工厂","podName":"武汉总代","podCarrierName":"捷达物流","carryDistance":662,"carryDays":2}]}
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            transport_path = Path(tmpdir) / "transport.html"
            sales_path = Path(tmpdir) / "sales.html"
            transport_path.write_text(transport_frontend, encoding="utf-8")
            sales_path.write_text(sales_frontend, encoding="utf-8")
            context = solve.load_frontend_contexts([transport_path, sales_path])

        assert context is not None
        self.assertEqual(context["source_kind"], "multi")
        self.assertEqual(len(context["carriers"]), 1)
        self.assertEqual(context["carriers"][0]["unit_price"], 0.31)
        self.assertEqual(context["segment_distances"]["苏州工厂-->北方总代"], 987.0)
        self.assertEqual(context["segment_distances"]["苏州工厂-->武汉总代"], 662.0)

    def test_heatwater_pattern_audit_is_opt_in(self) -> None:
        previous = __import__("os").environ.get("SUPPLY_CHAIN_HEATWATER_PATTERN_AUDIT")
        try:
            __import__("os").environ.pop("SUPPLY_CHAIN_HEATWATER_PATTERN_AUDIT", None)
            self.assertFalse(solve.heatwater_pattern_audit_enabled())
            __import__("os").environ["SUPPLY_CHAIN_HEATWATER_PATTERN_AUDIT"] = "1"
            self.assertTrue(solve.heatwater_pattern_audit_enabled())
        finally:
            if previous is None:
                __import__("os").environ.pop("SUPPLY_CHAIN_HEATWATER_PATTERN_AUDIT", None)
            else:
                __import__("os").environ["SUPPLY_CHAIN_HEATWATER_PATTERN_AUDIT"] = previous

    def test_heatwater_global_audit_is_opt_in(self) -> None:
        previous = __import__("os").environ.get("SUPPLY_CHAIN_HEATWATER_GLOBAL_AUDIT")
        try:
            __import__("os").environ.pop("SUPPLY_CHAIN_HEATWATER_GLOBAL_AUDIT", None)
            self.assertFalse(solve.heatwater_global_audit_enabled())
            __import__("os").environ["SUPPLY_CHAIN_HEATWATER_GLOBAL_AUDIT"] = "yes"
            self.assertTrue(solve.heatwater_global_audit_enabled())
        finally:
            if previous is None:
                __import__("os").environ.pop("SUPPLY_CHAIN_HEATWATER_GLOBAL_AUDIT", None)
            else:
                __import__("os").environ["SUPPLY_CHAIN_HEATWATER_GLOBAL_AUDIT"] = previous

    def test_heatwater_full_global_audit_is_opt_in(self) -> None:
        previous = __import__("os").environ.get("SUPPLY_CHAIN_HEATWATER_FULL_GLOBAL_AUDIT")
        try:
            __import__("os").environ.pop("SUPPLY_CHAIN_HEATWATER_FULL_GLOBAL_AUDIT", None)
            self.assertFalse(solve.heatwater_full_global_audit_enabled())
            __import__("os").environ["SUPPLY_CHAIN_HEATWATER_FULL_GLOBAL_AUDIT"] = "audit"
            self.assertTrue(solve.heatwater_full_global_audit_enabled())
        finally:
            if previous is None:
                __import__("os").environ.pop("SUPPLY_CHAIN_HEATWATER_FULL_GLOBAL_AUDIT", None)
            else:
                __import__("os").environ["SUPPLY_CHAIN_HEATWATER_FULL_GLOBAL_AUDIT"] = previous

    def test_concise_solver_status_reports_warm_window_audit(self) -> None:
        text = solve.concise_solver_status(
            {
                "warm_window_global_status": {
                    "status": "TimeLimitFeasibleNotSelected",
                    "warm_window_freight": 2165000.0,
                    "warm_base_freight": 2161114.0,
                    "solve_seconds": 180.1,
                    "time_limit_hit": True,
                    "model_stats": {
                        "lane_options": 33,
                        "shipment_vars": 1372,
                        "binary_vars": 1372,
                        "allowed_ship_day_option_windows": 11,
                        "allowed_ship_day_windows": 11,
                    },
                }
            }
        )

        self.assertIn("暖窗", text)
        self.assertIn("TimeLimitFeasibleNotSelected", text)
        self.assertIn("2,165,000", text)
        self.assertIn("模型lane33/x1372/bin1372/optwin11/routewin11/timelimit", text)

    def test_concise_solver_status_reports_nested_warm_window_audit(self) -> None:
        text = solve.concise_solver_status(
            {
                "product_transport": {
                    "warm_window_global_status": {
                        "status": "OptimalNoCostImprovement",
                        "warm_window_freight": 2162114.0,
                        "warm_base_freight": 2161114.0,
                        "strict_freight_upper_bound": 2161113.99,
                        "solve_seconds": 12.5,
                    }
                }
            }
        )

        self.assertIn("暖窗", text)
        self.assertIn("OptimalNoCostImprovement", text)
        self.assertIn("2,162,114", text)
        self.assertIn("反例<2,161,114", text)

    def test_concise_solver_status_reports_full_global_audit(self) -> None:
        text = solve.concise_solver_status(
            {
                "product_transport": {
                    "full_global_audit_status": {
                        "status": "TimeLimitFeasibleNotSelected",
                        "full_global_freight": 2161114.0,
                        "full_global_base_freight": 2161114.0,
                        "solve_seconds": 300.4,
                        "model_stats": {"lane_options": 33, "shipment_vars": 1768, "binary_vars": 1768},
                    }
                }
            }
        )

        self.assertIn("全路线审查", text)
        self.assertIn("TimeLimitFeasibleNotSelected", text)
        self.assertIn("2,161,114", text)
        self.assertIn("模型lane33/x1768/bin1768", text)

    def test_heatwater_extreme_inventory_window_handles_changed_numbers(self) -> None:
        previous_mode = solve.ACTIVE_SOLVER_MODE
        solve.set_solver_mode("extreme")
        try:
            days = 12
            products = [solve.Product("产品", "热水器", "台", bom=1.0, charge_ratio=0.5)]
            factories = [
                solve.Factory("东莞工厂", "热水器", "台", init=120, daily=24, limit=9999, excess_fee=0),
                solve.Factory("苏州工厂", "热水器", "台", init=90, daily=30, limit=9999, excess_fee=0),
            ]
            forecasts = [
                {
                    "node": "武汉总代",
                    "product": "热水器",
                    "forecast": 168,
                    "daily_avg": 14,
                    "daily_std": 0,
                    "safety_days": 1,
                    "safety_stock": 14,
                    "init": 40,
                    "limit": 70,
                    "excess_fee": 5,
                    "daily_demand": [14] * days,
                    "plan_days": days,
                },
                {
                    "node": "北方总代",
                    "product": "热水器",
                    "forecast": 204,
                    "daily_avg": 17,
                    "daily_std": 0,
                    "safety_days": 1,
                    "safety_stock": 17,
                    "init": 55,
                    "limit": 85,
                    "excess_fee": 5,
                    "daily_demand": [17] * days,
                    "plan_days": days,
                },
            ]
            carriers = [
                Carrier("捷达物流", "汽运", 600.0, 0.35, 12000.0),
                Carrier("顺达物流", "汽运", 460.0, 0.45, 8000.0),
                Carrier("中铁快运", "铁路", 300.0, 0.11, 30000.0),
            ]
            routes = [
                solve.Route("东莞工厂-->武汉总代", "东莞工厂", "武汉总代", 820.0, 0.0, 0.0, 0.0, 2, "CNY"),
                solve.Route("苏州工厂-->武汉总代", "苏州工厂", "武汉总代", 660.0, 0.0, 0.0, 0.0, 2, "CNY"),
                solve.Route("东莞工厂-->北方总代", "东莞工厂", "北方总代", 1200.0, 0.0, 0.0, 0.0, 3, "CNY"),
                solve.Route("苏州工厂-->北方总代", "苏州工厂", "北方总代", 980.0, 0.0, 0.0, 0.0, 2, "CNY"),
            ]

            rows, status = solve.build_inventory_window_extreme_sales_transport(
                forecasts=forecasts,
                factories=factories,
                routes=routes,
                products=products,
                cargo="热水器",
                carriers=carriers,
                days=days,
            )

            replay = solve.daily_market_replay(forecasts, rows, days)
            hard_risks = solve.product_transport_hard_risks(
                forecasts=forecasts,
                shipments=rows,
                factories=factories,
                days=days,
            )
            self.assertEqual(status["status"], "OptimalWindowCandidate")
            self.assertTrue(solve.is_full_satisfaction(replay["market_satisfaction"], replay["shortage"]))
            self.assertEqual(hard_risks, [])
            self.assertTrue(all(isinstance(row["amount"], int) and row["amount"] > 0 for row in rows))
        finally:
            solve.set_solver_mode(previous_mode)

    def test_day_transport_milp_honors_option_level_day_windows(self) -> None:
        previous_mode = solve.ACTIVE_SOLVER_MODE
        solve.set_solver_mode("extreme")
        try:
            products = [solve.Product("产品", "热水器", "台", bom=1.0, charge_ratio=1.0)]
            carriers = [
                Carrier("便宜慢车", "汽运", 1.0, 1.0, 0.0),
                Carrier("昂贵快车", "汽运", 10.0, 10.0, 0.0),
            ]
            route = solve.Route("F-->A", "F", "A", 10.0, 0.0, 0.0, 0.0, 1, "CNY")
            transport = solve._solve_day_transport_milp(
                name="option_window_unit",
                sources={"F": {"initial": 20, "supply": [0, 0], "supply_is_capacity": False, "limit": 0}},
                destinations={"A": {"initial": 0, "demand": [0, 10], "limit": 20, "excess_fee": 0}},
                routes=[route],
                products=products,
                cargo="热水器",
                carriers=carriers,
                days=2,
                enforce_destination_limits=True,
                max_total_shortage=0.0,
                allowed_ship_days_by_option={("F-->A", "F", "A", "昂贵快车"): {1}},
                prune_transport_options=False,
            )

            self.assertEqual(transport["status"], "Optimal")
            self.assertEqual(transport["shipments"][0]["carrier"], "昂贵快车")
            self.assertEqual(transport["model_stats"]["allowed_ship_day_option_windows"], 1)
            self.assertEqual(transport["model_stats"]["shipment_vars"], 1)
        finally:
            solve.set_solver_mode(previous_mode)

    def test_day_transport_milp_strict_freight_bound_does_not_fallback(self) -> None:
        previous_mode = solve.ACTIVE_SOLVER_MODE
        solve.set_solver_mode("extreme")
        try:
            products = [solve.Product("产品", "热水器", "台", bom=1.0, charge_ratio=1.0)]
            carriers = [Carrier("便宜车", "汽运", 100.0, 1.0, 0.0)]
            route = solve.Route("F-->A", "F", "A", 1.0, 0.0, 0.0, 0.0, 0, "CNY")
            transport = solve._solve_day_transport_milp(
                name="strict_bound_no_fallback_unit",
                sources={"F": {"initial": 10, "supply": [0], "supply_is_capacity": False, "limit": 0}},
                destinations={"A": {"initial": 0, "demand": [10], "limit": 10, "excess_fee": 0}},
                routes=[route],
                products=products,
                cargo="热水器",
                carriers=carriers,
                days=1,
                enforce_destination_limits=True,
                max_total_shortage=0.0,
                strict_freight_upper_bound=9.99,
                allow_fallback=False,
            )

            self.assertEqual(transport["status"], "Infeasible")
            self.assertEqual(transport["shipments"], [])
            self.assertEqual(transport["strict_freight_upper_bound"], 9.99)
        finally:
            solve.set_solver_mode(previous_mode)

    def test_day_transport_milp_strict_feasibility_reports_platform_freight(self) -> None:
        previous_mode = solve.ACTIVE_SOLVER_MODE
        solve.set_solver_mode("extreme")
        try:
            products = [solve.Product("产品", "热水器", "台", bom=1.0, charge_ratio=1.0)]
            carriers = [Carrier("平台车", "汽运", 100.0, 1.0, 0.0)]
            route = solve.Route("F-->A", "F", "A", 10.0, 0.0, 0.0, 0.0, 0, "CNY")
            transport = solve._solve_day_transport_milp(
                name="strict_bound_feasibility_platform_cost_unit",
                sources={"F": {"initial": 10, "supply": [0, 0], "supply_is_capacity": False, "limit": 0}},
                destinations={"A": {"initial": 0, "demand": [0, 10], "limit": 10, "excess_fee": 0}},
                routes=[route],
                products=products,
                cargo="热水器",
                carriers=carriers,
                days=2,
                enforce_destination_limits=True,
                max_total_shortage=0.0,
                strict_freight_upper_bound=500.0,
                allow_fallback=False,
                strict_bound_feasibility_only=True,
            )

            self.assertEqual(transport["status"], "Optimal")
            self.assertEqual(transport["model_stats"]["strict_bound_feasibility_only"], True)
            self.assertEqual(len(transport["shipments"]), 1)
            self.assertEqual(transport["shipments"][0]["freight_cost"], 100.0)
            self.assertEqual(transport["freight_cost"], 100.0)
        finally:
            solve.set_solver_mode(previous_mode)

    def test_fixed_arrival_reroute_finds_cheaper_lane_without_changing_arrival(self) -> None:
        previous_mode = solve.ACTIVE_SOLVER_MODE
        solve.set_solver_mode("extreme")
        try:
            products = [solve.Product("产品", "热水器", "台", bom=1.0, charge_ratio=1.0)]
            factories = [
                solve.Factory("贵厂", "热水器", "台", init=100, daily=0, limit=9999, excess_fee=0),
                solve.Factory("便宜厂", "热水器", "台", init=100, daily=0, limit=9999, excess_fee=0),
            ]
            routes = [
                solve.Route("贵厂-->A门店", "贵厂", "A门店", 10.0, 0.0, 0.0, 0.0, 1, "CNY"),
                solve.Route("便宜厂-->A门店", "便宜厂", "A门店", 2.0, 0.0, 0.0, 0.0, 1, "CNY"),
            ]
            carriers = [Carrier("汽运", "汽运", 10.0, 1.0, 0.0)]
            forecasts = [{"node": "A门店", "init": 0, "limit": 100, "daily_demand": [0, 10], "forecast": 10}]
            shipments = [
                {
                    "destination": "A门店",
                    "factory": "贵厂",
                    "source": "贵厂",
                    "cargo": "热水器",
                    "amount": 10,
                    "route": "贵厂-->A门店",
                    "carrier": "汽运",
                    "lead": 1,
                    "ship_day": 1,
                    "arrival_day": 2,
                    "freight_cost": 100.0,
                }
            ]

            rows, status = solve.optimize_fixed_arrival_reroute(
                forecasts=forecasts,
                shipments=shipments,
                factories=factories,
                routes=routes,
                products=products,
                cargo="热水器",
                carriers=carriers,
                days=2,
            )

            self.assertEqual(status["status"], "Optimal")
            self.assertGreater(status["delta"], 0)
            self.assertEqual(rows[0]["source"], "便宜厂")
            self.assertEqual(rows[0]["arrival_day"], 2)
        finally:
            solve.set_solver_mode(previous_mode)

    def test_best_arrival_option_counts_future_local_factory_usage(self) -> None:
        products = [solve.Product("产品", "热水器", "台", bom=1.0, charge_ratio=1.0)]
        factories = [
            solve.Factory("便宜厂", "热水器", "台", init=10, daily=0, limit=9999, excess_fee=0),
            solve.Factory("备用厂", "热水器", "台", init=20, daily=0, limit=9999, excess_fee=0),
        ]
        routes = [
            solve.Route("便宜厂-->A门店", "便宜厂", "A门店", 1.0, 1.0, 0.0, 0.0, 0, "CNY"),
            solve.Route("备用厂-->A门店", "备用厂", "A门店", 1.0, 5.0, 0.0, 0.0, 0, "CNY"),
        ]
        output_by_factory = solve.factory_output_schedule(factories, 3, None)
        capacity_prefix = solve._factory_capacity_prefix(factories, output_by_factory, 3)
        local_rows = [
            {
                "destination": "B门店",
                "factory": "便宜厂",
                "source": "便宜厂",
                "cargo": "热水器",
                "amount": 6,
                "route": "便宜厂-->B门店",
                "ship_day": 3,
                "arrival_day": 3,
                "freight_cost": 6.0,
            }
        ]

        row = solve._best_arrival_option_row(
            node="A门店",
            amount=6,
            arrival_day=1,
            factories=factories,
            routes=routes,
            products=products,
            cargo="热水器",
            carriers=None,
            days=3,
            base_rows=[],
            local_rows=local_rows,
            output_by_factory=output_by_factory,
            capacity_prefix=capacity_prefix,
        )

        self.assertIsNotNone(row)
        self.assertEqual(row["source"], "备用厂")


    def test_destination_block_polish_reuses_arrival_option_cache(self) -> None:
        products = [solve.Product("产品", "热水器", "台", bom=1.0, charge_ratio=1.0)]
        factories = [solve.Factory("F", "热水器", "台", init=300, daily=0, limit=9999, excess_fee=0)]
        forecasts = [
            {
                "node": node,
                "product": "热水器",
                "forecast": 20,
                "init": 0,
                "limit": 20,
                "excess_fee": 0,
                "daily_demand": [10, 10],
                "plan_days": 2,
            }
            for node in ("A", "B", "C")
        ]
        routes = [solve.Route(f"F-->{node}", "F", node, 1.0, 1.0, 0.0, 0.0, 0, "CNY") for node in ("A", "B", "C")]
        shipments = [
            {
                "destination": node,
                "source": "F",
                "factory": "F",
                "cargo": "热水器",
                "amount": 20,
                "route": f"F-->{node}",
                "lead": 0,
                "ship_day": 1,
                "arrival_day": 1,
                "freight_cost": 20.0,
            }
            for node in ("A", "B", "C")
        ]
        original = solve.route_transport_options
        calls = {"count": 0}

        def counted_route_transport_options(*args, **kwargs):
            calls["count"] += 1
            return original(*args, **kwargs)

        solve.route_transport_options = counted_route_transport_options
        try:
            _rows, status = solve.polish_destination_blocks(
                forecasts=forecasts,
                shipments=shipments,
                factories=factories,
                routes=routes,
                products=products,
                cargo="热水器",
                carriers=None,
                days=2,
                max_block_size=3,
                max_rounds=1,
            )
        finally:
            solve.route_transport_options = original

        self.assertEqual(status["option_cache_entries"], 18)
        self.assertLessEqual(calls["count"], status["option_cache_entries"])

    def test_shift_polish_moves_surplus_from_expensive_early_lane(self) -> None:
        products = [solve.Product("产品", "热水器", "台", bom=1.0, charge_ratio=1.0)]
        factories = [solve.Factory("F", "热水器", "台", init=1000, daily=0, limit=9999, excess_fee=0)]
        routes = [
            solve.Route("F-->A", "F", "A", 1.0, 10.0, 0.0, 0.0, 0, "CNY"),
            solve.Route("F-->X-->A", "F", "A", 1.0, 1.0, 0.0, 0.0, 4, "CNY"),
        ]
        forecasts = [
            {
                "node": "A",
                "product": "热水器",
                "forecast": 100,
                "init": 50,
                "limit": 200,
                "excess_fee": 0,
                "daily_demand": [10] * 10,
                "plan_days": 10,
            }
        ]
        shipments = [
            {
                "destination": "A",
                "source": "F",
                "factory": "F",
                "cargo": "热水器",
                "amount": 70,
                "route": "F-->A",
                "lead": 0,
                "ship_day": 1,
                "arrival_day": 1,
                "freight_cost": 700.0,
            },
            {
                "destination": "A",
                "source": "F",
                "factory": "F",
                "cargo": "热水器",
                "amount": 20,
                "route": "F-->X-->A",
                "lead": 4,
                "ship_day": 1,
                "arrival_day": 5,
                "freight_cost": 20.0,
            },
        ]

        polished, status = solve.polish_shift_expensive_early_to_later(
            forecasts=forecasts,
            shipments=shipments,
            factories=factories,
            routes=routes,
            products=products,
            cargo="热水器",
            carriers=None,
            days=10,
        )

        self.assertGreater(status["delta"], 0)
        self.assertEqual(solve.daily_market_replay(forecasts, polished, 10)["shortage"], 0)
        self.assertLess(sum(row["freight_cost"] for row in polished), 720.0)

    def test_shift_polish_moves_expensive_late_lane_into_earlier_capacity(self) -> None:
        products = [solve.Product("产品", "热水器", "台", bom=1.0, charge_ratio=1.0)]
        factories = [solve.Factory("F", "热水器", "台", init=1000, daily=0, limit=9999, excess_fee=0)]
        routes = [
            solve.Route("F-->A", "F", "A", 1.0, 1.0, 0.0, 0.0, 0, "CNY"),
            solve.Route("F-->Y-->A", "F", "A", 1.0, 10.0, 0.0, 0.0, 4, "CNY"),
        ]
        forecasts = [
            {
                "node": "A",
                "product": "热水器",
                "forecast": 100,
                "init": 50,
                "limit": 200,
                "excess_fee": 0,
                "daily_demand": [10] * 10,
                "plan_days": 10,
            }
        ]
        shipments = [
            {
                "destination": "A",
                "source": "F",
                "factory": "F",
                "cargo": "热水器",
                "amount": 20,
                "route": "F-->A",
                "lead": 0,
                "ship_day": 5,
                "arrival_day": 5,
                "freight_cost": 20.0,
            },
            {
                "destination": "A",
                "source": "F",
                "factory": "F",
                "cargo": "热水器",
                "amount": 50,
                "route": "F-->Y-->A",
                "lead": 4,
                "ship_day": 6,
                "arrival_day": 10,
                "freight_cost": 500.0,
            },
        ]

        polished, status = solve.polish_shift_expensive_early_to_later(
            forecasts=forecasts,
            shipments=shipments,
            factories=factories,
            routes=routes,
            products=products,
            cargo="热水器",
            carriers=None,
            days=10,
        )

        self.assertGreater(status["delta"], 0)
        self.assertEqual(solve.daily_market_replay(forecasts, polished, 10)["shortage"], 0)
        self.assertLess(sum(row["freight_cost"] for row in polished), 520.0)

    def test_shift_polish_merges_equal_unit_minimum_charge_shipments(self) -> None:
        products = [solve.Product("产品", "热水器", "台", bom=1.0, charge_ratio=1.0)]
        factories = [solve.Factory("F", "热水器", "台", init=1000, daily=0, limit=9999, excess_fee=0)]
        carriers = [Carrier("最低费车", "汽运", 100.0, 0.01, 100.0)]
        routes = [solve.Route("F-->A", "F", "A", 1.0, 0.0, 0.0, 0.0, 0, "CNY")]
        forecasts = [
            {
                "node": "A",
                "product": "热水器",
                "forecast": 20,
                "init": 10,
                "limit": 30,
                "excess_fee": 0,
                "daily_demand": [1] * 20,
                "plan_days": 20,
            }
        ]
        shipments = [
            {
                "destination": "A",
                "source": "F",
                "factory": "F",
                "cargo": "热水器",
                "amount": 10,
                "route": "F-->A",
                "carrier": "最低费车",
                "lead": 1,
                "ship_day": 1,
                "arrival_day": 2,
                "freight_cost": 100.0,
            },
            {
                "destination": "A",
                "source": "F",
                "factory": "F",
                "cargo": "热水器",
                "amount": 10,
                "route": "F-->A",
                "carrier": "最低费车",
                "lead": 1,
                "ship_day": 6,
                "arrival_day": 7,
                "freight_cost": 100.0,
            },
        ]

        polished, status = solve.polish_shift_expensive_early_to_later(
            forecasts=forecasts,
            shipments=shipments,
            factories=factories,
            routes=routes,
            products=products,
            cargo="热水器",
            carriers=carriers,
            days=20,
        )

        self.assertGreater(status["delta"], 0)
        self.assertEqual(sum(row["freight_cost"] for row in polished), 100.0)
        self.assertEqual(solve.daily_market_replay(forecasts, polished, 20)["shortage"], 0)

    def test_shift_polish_rejects_cheaper_merge_that_causes_stockout(self) -> None:
        products = [solve.Product("产品", "热水器", "台", bom=1.0, charge_ratio=1.0)]
        factories = [solve.Factory("F", "热水器", "台", init=1000, daily=0, limit=9999, excess_fee=0)]
        carriers = [Carrier("最低费车", "汽运", 100.0, 0.01, 100.0)]
        routes = [solve.Route("F-->A", "F", "A", 1.0, 0.0, 0.0, 0.0, 0, "CNY")]
        forecasts = [
            {
                "node": "A",
                "product": "热水器",
                "forecast": 20,
                "init": 0,
                "limit": 30,
                "excess_fee": 0,
                "daily_demand": [2] * 10,
                "plan_days": 10,
            }
        ]
        shipments = [
            {
                "destination": "A",
                "source": "F",
                "factory": "F",
                "cargo": "热水器",
                "amount": 10,
                "route": "F-->A",
                "carrier": "最低费车",
                "lead": 1,
                "ship_day": 1,
                "arrival_day": 2,
                "freight_cost": 100.0,
            },
            {
                "destination": "A",
                "source": "F",
                "factory": "F",
                "cargo": "热水器",
                "amount": 10,
                "route": "F-->A",
                "carrier": "最低费车",
                "lead": 1,
                "ship_day": 6,
                "arrival_day": 7,
                "freight_cost": 100.0,
            },
        ]

        polished, status = solve.polish_shift_expensive_early_to_later(
            forecasts=forecasts,
            shipments=shipments,
            factories=factories,
            routes=routes,
            products=products,
            cargo="热水器",
            carriers=carriers,
            days=10,
        )

        self.assertEqual(status["delta"], 0.0)
        self.assertEqual(sum(row["freight_cost"] for row in polished), 200.0)
        self.assertEqual(solve.daily_market_replay(forecasts, polished, 10)["shortage"], 2)

    def test_delay_polish_replaces_urgent_direct_lane_when_inventory_can_wait(self) -> None:
        products = [solve.Product("产品", "热水器", "台", bom=1.0, charge_ratio=1.0)]
        factories = [solve.Factory("F", "热水器", "台", init=1000, daily=0, limit=9999, excess_fee=0)]
        routes = [
            solve.Route("F-->A", "F", "A", 1.0, 10.0, 0.0, 0.0, 0, "CNY"),
            solve.Route("F-->X-->A", "F", "A", 1.0, 1.0, 0.0, 0.0, 4, "CNY"),
        ]
        forecasts = [
            {
                "node": "A",
                "product": "热水器",
                "forecast": 50,
                "init": 30,
                "limit": 100,
                "excess_fee": 0,
                "daily_demand": [5] * 10,
                "plan_days": 10,
            }
        ]
        shipments = [
            {
                "destination": "A",
                "source": "F",
                "factory": "F",
                "cargo": "热水器",
                "amount": 50,
                "route": "F-->A",
                "carrier": "汽运承运商",
                "lead": 0,
                "ship_day": 1,
                "arrival_day": 1,
                "freight_cost": 500.0,
            }
        ]

        polished, status = solve.polish_delay_to_cheaper_lanes(
            forecasts=forecasts,
            shipments=shipments,
            factories=factories,
            routes=routes,
            products=products,
            cargo="热水器",
            carriers=None,
            days=10,
        )

        self.assertGreater(status["delta"], 0)
        self.assertEqual(polished[0]["route"], "F-->X-->A")
        self.assertEqual(polished[0]["arrival_day"], 5)
        self.assertEqual(sum(row["freight_cost"] for row in polished), 50.0)
        self.assertEqual(solve.daily_market_replay(forecasts, polished, 10)["shortage"], 0)

    def test_heatwater_gap_diagnostics_flags_independent_capacity_conflict(self) -> None:
        forecasts = [
            {"node": "A", "daily_demand": [5, 5], "init": 0, "limit": 20},
            {"node": "B", "daily_demand": [5, 5], "init": 0, "limit": 20},
        ]
        factories = [solve.Factory("苏州工厂", "热水器", "台", init=0, daily=5, limit=9999, excess_fee=0)]
        independent = {
            "status": "Computed",
            "freight_lower_bound": 200.0,
            "node_costs": {"A": 100.0, "B": 100.0},
            "node_shipments": {
                "A": [{"destination": "A", "source": "苏州工厂", "amount": 10, "ship_day": 1, "freight_cost": 100.0}],
                "B": [{"destination": "B", "source": "苏州工厂", "amount": 10, "ship_day": 1, "freight_cost": 100.0}],
            },
        }
        final_rows = [
            {"destination": "A", "source": "苏州工厂", "amount": 10, "ship_day": 1, "freight_cost": 100.0},
            {"destination": "B", "source": "苏州工厂", "amount": 10, "ship_day": 2, "freight_cost": 130.0},
        ]

        diagnostics = solve.heatwater_gap_diagnostics(
            forecasts=forecasts,
            factories=factories,
            final_shipments=final_rows,
            independent_status=independent,
            days=2,
        )

        self.assertEqual(diagnostics["status"], "Computed")
        self.assertFalse(diagnostics["independent_capacity_feasible"])
        self.assertEqual(diagnostics["max_independent_capacity_overflow"], 15)
        self.assertEqual(diagnostics["node_deltas"]["B"]["delta"], 30.0)
        text = solve.concise_solver_status(
            {
                "lower_bound_status": {
                    "status": "Computed",
                    "feasible_freight": 230.0,
                    "freight_lower_bound": 200.0,
                    "gap_ratio": 30.0 / 230.0,
                    "gap_diagnostics": diagnostics,
                }
            }
        )
        self.assertIn("独立下界产能审查：不可行", text)

    def test_pattern_combinations_record_conflict_shadow_patterns(self) -> None:
        products = [solve.Product("产品", "热水器", "台", bom=1.0, charge_ratio=1.0)]
        factories = [
            solve.Factory("F", "热水器", "台", init=0, daily=10, limit=9999, excess_fee=0),
            solve.Factory("G", "热水器", "台", init=100, daily=0, limit=9999, excess_fee=0),
        ]
        forecasts = [
            {"node": "A", "daily_demand": [10, 10], "init": 0, "limit": 30},
            {"node": "B", "daily_demand": [10, 10], "init": 0, "limit": 30},
        ]
        routes = [
            solve.Route("F-->A", "F", "A", 1.0, 1.0, 0.0, 0.0, 0, "CNY"),
            solve.Route("F-->B", "F", "B", 1.0, 1.0, 0.0, 0.0, 0, "CNY"),
            solve.Route("G-->A", "G", "A", 1.0, 3.0, 0.0, 0.0, 0, "CNY"),
            solve.Route("G-->B", "G", "B", 1.0, 3.0, 0.0, 0.0, 0, "CNY"),
        ]
        shipments = [
            {"destination": "A", "source": "F", "factory": "F", "cargo": "热水器", "amount": 20, "route": "F-->A", "carrier": "汽运承运商", "lead": 0, "ship_day": 1, "arrival_day": 1, "freight_cost": 20.0},
            {"destination": "B", "source": "G", "factory": "G", "cargo": "热水器", "amount": 20, "route": "G-->B", "carrier": "汽运承运商", "lead": 0, "ship_day": 1, "arrival_day": 1, "freight_cost": 60.0},
        ]

        _rows, status = solve.polish_destination_pattern_combinations(
            forecasts=forecasts,
            shipments=shipments,
            factories=factories,
            routes=routes,
            products=products,
            cargo="热水器",
            carriers=None,
            days=2,
            max_patterns_per_destination=8,
        )

        self.assertIn("conflict_shadow_patterns", status)
        self.assertGreaterEqual(status["conflict_shadow_patterns"], 0)
        self.assertIn("pattern_milp", status)

    def test_pattern_audit_records_factory_avoidance_patterns(self) -> None:
        previous = os.environ.get("SUPPLY_CHAIN_HEATWATER_PATTERN_AUDIT")
        os.environ["SUPPLY_CHAIN_HEATWATER_PATTERN_AUDIT"] = "1"
        try:
            products = [solve.Product("产品", "热水器", "台", bom=1.0, charge_ratio=1.0)]
            factories = [
                solve.Factory("拥堵厂", "热水器", "台", init=50, daily=10, limit=9999, excess_fee=0),
                solve.Factory("备用厂", "热水器", "台", init=50, daily=10, limit=9999, excess_fee=0),
            ]
            forecasts = [
                {"node": "A", "daily_demand": [5, 5], "init": 0, "limit": 20},
                {"node": "B", "daily_demand": [5, 5], "init": 0, "limit": 20},
            ]
            routes = [
                solve.Route("拥堵厂-->A", "拥堵厂", "A", 1.0, 1.0, 0.0, 0.0, 0, "CNY"),
                solve.Route("备用厂-->A", "备用厂", "A", 1.0, 2.0, 0.0, 0.0, 0, "CNY"),
                solve.Route("拥堵厂-->B", "拥堵厂", "B", 1.0, 1.0, 0.0, 0.0, 0, "CNY"),
                solve.Route("备用厂-->B", "备用厂", "B", 1.0, 2.0, 0.0, 0.0, 0, "CNY"),
            ]
            shipments = [
                {"destination": "A", "source": "拥堵厂", "factory": "拥堵厂", "cargo": "热水器", "amount": 10, "route": "拥堵厂-->A", "carrier": "汽运承运商", "lead": 0, "ship_day": 1, "arrival_day": 1, "freight_cost": 10.0},
                {"destination": "B", "source": "拥堵厂", "factory": "拥堵厂", "cargo": "热水器", "amount": 10, "route": "拥堵厂-->B", "carrier": "汽运承运商", "lead": 0, "ship_day": 1, "arrival_day": 1, "freight_cost": 10.0},
            ]

            _rows, status = solve.polish_destination_pattern_combinations(
                forecasts=forecasts,
                shipments=shipments,
                factories=factories,
                routes=routes,
                products=products,
                cargo="热水器",
                carriers=None,
                days=2,
                max_patterns_per_destination=8,
            )

            self.assertIn("avoidance_patterns", status)
            self.assertGreaterEqual(status["avoidance_patterns"], 0)
        finally:
            if previous is None:
                os.environ.pop("SUPPLY_CHAIN_HEATWATER_PATTERN_AUDIT", None)
            else:
                os.environ["SUPPLY_CHAIN_HEATWATER_PATTERN_AUDIT"] = previous


if __name__ == "__main__":
    unittest.main()
