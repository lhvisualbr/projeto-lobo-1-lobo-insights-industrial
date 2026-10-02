from __future__ import annotations

import sys
import unittest
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


from intelligence.cost_rules import (  # noqa: E402
    evaluate_cost_rules,
    rule_cost_concentration,
    rule_high_cost_and_consumption_impact,
)
from intelligence.models import Severity  # noqa: E402


class CostRulesTests(unittest.TestCase):
    def make_row(
        self,
        *,
        material_id: str = "MAT001",
        material: str = "Material Teste",
        cost: float = 2500.0,
        cost_share: float = 12.0,
        consumption_share: float = 8.0,
    ) -> pd.Series:
        return pd.Series(
            {
                "Material_ID": material_id,
                "Material": material,
                "custo_consumo": cost,
                "participacao_custo_pct": cost_share,
                "participacao_consumo_pct": consumption_share,
            }
        )

    def test_cost_concentration_creates_attention_signal(
        self,
    ) -> None:
        row = self.make_row(
            cost_share=12.0,
        )

        signal = rule_cost_concentration(
            row
        )

        self.assertIsNotNone(
            signal
        )
        assert signal is not None

        self.assertEqual(
            signal.id,
            "COST-CONCENTRATION-MAT001",
        )

        self.assertEqual(
            signal.type,
            "CONCENTRACAO_CUSTO",
        )

        self.assertEqual(
            signal.severity,
            Severity.ATENCAO,
        )

        self.assertEqual(
            signal.entity,
            "MAT001",
        )

        self.assertEqual(
            signal.rule_origin,
            "rule_cost_concentration",
        )

        self.assertEqual(
            len(signal.evidences),
            2,
        )

    def test_cost_concentration_becomes_high_at_15_percent(
        self,
    ) -> None:
        row = self.make_row(
            cost_share=15.0,
        )

        signal = rule_cost_concentration(
            row
        )

        self.assertIsNotNone(
            signal
        )
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.ALTO,
        )

    def test_cost_concentration_accepts_exact_10_percent(
        self,
    ) -> None:
        row = self.make_row(
            cost_share=10.0,
        )

        signal = rule_cost_concentration(
            row
        )

        self.assertIsNotNone(
            signal
        )
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.ATENCAO,
        )

    def test_cost_concentration_ignores_below_threshold(
        self,
    ) -> None:
        row = self.make_row(
            cost_share=9.99,
        )

        signal = rule_cost_concentration(
            row
        )

        self.assertIsNone(
            signal
        )

    def test_cost_concentration_ignores_zero_cost(
        self,
    ) -> None:
        row = self.make_row(
            cost=0.0,
            cost_share=20.0,
        )

        signal = rule_cost_concentration(
            row
        )

        self.assertIsNone(
            signal
        )

    def test_combined_impact_creates_high_signal(
        self,
    ) -> None:
        row = self.make_row(
            cost_share=12.0,
            consumption_share=14.0,
        )

        signal = (
            rule_high_cost_and_consumption_impact(
                row
            )
        )

        self.assertIsNotNone(
            signal
        )
        assert signal is not None

        self.assertEqual(
            signal.id,
            "COST-CONSUMPTION-IMPACT-MAT001",
        )

        self.assertEqual(
            signal.type,
            "ALTO_IMPACTO_CUSTO_CONSUMO",
        )

        self.assertEqual(
            signal.severity,
            Severity.ALTO,
        )

        self.assertEqual(
            signal.entity,
            "MAT001",
        )

        self.assertEqual(
            signal.rule_origin,
            "rule_high_cost_and_consumption_impact",
        )

        self.assertEqual(
            len(signal.evidences),
            3,
        )

    def test_combined_impact_accepts_exact_thresholds(
        self,
    ) -> None:
        row = self.make_row(
            cost_share=10.0,
            consumption_share=10.0,
        )

        signal = (
            rule_high_cost_and_consumption_impact(
                row
            )
        )

        self.assertIsNotNone(
            signal
        )

    def test_combined_impact_ignores_low_cost_share(
        self,
    ) -> None:
        row = self.make_row(
            cost_share=9.0,
            consumption_share=20.0,
        )

        signal = (
            rule_high_cost_and_consumption_impact(
                row
            )
        )

        self.assertIsNone(
            signal
        )

    def test_combined_impact_ignores_low_consumption_share(
        self,
    ) -> None:
        row = self.make_row(
            cost_share=20.0,
            consumption_share=9.0,
        )

        signal = (
            rule_high_cost_and_consumption_impact(
                row
            )
        )

        self.assertIsNone(
            signal
        )

    def test_invalid_numeric_value_is_ignored(
        self,
    ) -> None:
        row = self.make_row()

        row[
            "participacao_custo_pct"
        ] = "invalido"

        self.assertIsNone(
            rule_cost_concentration(
                row
            )
        )

    def test_nan_value_is_ignored(
        self,
    ) -> None:
        row = self.make_row(
            cost_share=float("nan"),
        )

        self.assertIsNone(
            rule_cost_concentration(
                row
            )
        )

    def test_evaluator_returns_deterministic_order(
        self,
    ) -> None:
        features = pd.DataFrame(
            [
                {
                    "Material_ID": "MAT003",
                    "Material": "Material C",
                    "custo_consumo": 2000.0,
                    "participacao_custo_pct": 12.0,
                    "participacao_consumo_pct": 5.0,
                },
                {
                    "Material_ID": "MAT002",
                    "Material": "Material B",
                    "custo_consumo": 4000.0,
                    "participacao_custo_pct": 16.0,
                    "participacao_consumo_pct": 15.0,
                },
                {
                    "Material_ID": "MAT001",
                    "Material": "Material A",
                    "custo_consumo": 3000.0,
                    "participacao_custo_pct": 11.0,
                    "participacao_consumo_pct": 12.0,
                },
            ]
        )

        signals = evaluate_cost_rules(
            features
        )

        self.assertEqual(
            len(signals),
            5,
        )

        self.assertEqual(
            signals[0].id,
            "COST-CONCENTRATION-MAT002",
        )

        self.assertEqual(
            signals[1].id,
            "COST-CONSUMPTION-IMPACT-MAT001",
        )

        self.assertEqual(
            signals[2].id,
            "COST-CONSUMPTION-IMPACT-MAT002",
        )

        self.assertEqual(
            signals[3].id,
            "COST-CONCENTRATION-MAT001",
        )

        self.assertEqual(
            signals[4].id,
            "COST-CONCENTRATION-MAT003",
        )

    def test_evaluator_rejects_missing_columns(
        self,
    ) -> None:
        features = pd.DataFrame(
            [
                {
                    "Material_ID": "MAT001",
                    "Material": "Material A",
                }
            ]
        )

        with self.assertRaises(
            ValueError
        ):
            evaluate_cost_rules(
                features
            )


if __name__ == "__main__":
    unittest.main()