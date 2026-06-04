from __future__ import annotations

import score_replay
import solve
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


if __name__ == "__main__":
    unittest.main()
