from __future__ import annotations

import unittest

import pandas as pd

from src.intelligence.consumption_rules import (
    evaluate_consumption_rules,
    rule_consumption_drop,
    rule_consumption_growth,
)
from src.intelligence.models import Severity


class ConsumptionRulesTests(unittest.TestCase):
    def make_row(
        self,
        *,
        material_id: str = "MAT001",
        material: str = "Material Teste",
        previous: float = 100.0,
        latest: float = 140.0,
        variation: float = 40.0,
    ) -> pd.Series:
        return pd.Series(
            {
                "Material_ID": material_id,
                "Material": material,
                "consumo_mes_anterior": previous,
                "consumo_ultimo_mes": latest,
                "variacao_consumo_pct": variation,
            }
        )

    def test_growth_creates_attention_signal(self) -> None:
        row = self.make_row(
            previous=100.0,
            latest=140.0,
            variation=40.0,
        )

        signal = rule_consumption_growth(row)

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.id,
            "CONSUMPTION-GROWTH-MAT001",
        )
        self.assertEqual(
            signal.type,
            "CRESCIMENTO_CONSUMO",
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
            "rule_consumption_growth",
        )
        self.assertEqual(
            len(signal.evidences),
            3,
        )

    def test_growth_becomes_high_at_75_percent(self) -> None:
        row = self.make_row(
            previous=100.0,
            latest=175.0,
            variation=75.0,
        )

        signal = rule_consumption_growth(row)

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.ALTO,
        )

    def test_growth_ignores_variation_below_threshold(self) -> None:
        row = self.make_row(
            previous=100.0,
            latest=129.0,
            variation=29.0,
        )

        signal = rule_consumption_growth(row)

        self.assertIsNone(signal)

    def test_growth_ignores_zero_previous_consumption(self) -> None:
        row = self.make_row(
            previous=0.0,
            latest=50.0,
            variation=100.0,
        )

        signal = rule_consumption_growth(row)

        self.assertIsNone(signal)

    def test_growth_ignores_inconsistent_direction(self) -> None:
        row = self.make_row(
            previous=100.0,
            latest=90.0,
            variation=40.0,
        )

        signal = rule_consumption_growth(row)

        self.assertIsNone(signal)

    def test_drop_creates_info_signal(self) -> None:
        row = self.make_row(
            previous=100.0,
            latest=65.0,
            variation=-35.0,
        )

        signal = rule_consumption_drop(row)

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.id,
            "CONSUMPTION-DROP-MAT001",
        )
        self.assertEqual(
            signal.type,
            "REDUCAO_CONSUMO",
        )
        self.assertEqual(
            signal.severity,
            Severity.INFO,
        )
        self.assertEqual(
            signal.entity,
            "MAT001",
        )
        self.assertEqual(
            signal.rule_origin,
            "rule_consumption_drop",
        )
        self.assertEqual(
            len(signal.evidences),
            3,
        )

    def test_drop_becomes_attention_at_minus_50_percent(self) -> None:
        row = self.make_row(
            previous=100.0,
            latest=50.0,
            variation=-50.0,
        )

        signal = rule_consumption_drop(row)

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.ATENCAO,
        )

    def test_drop_ignores_variation_above_threshold(self) -> None:
        row = self.make_row(
            previous=100.0,
            latest=80.0,
            variation=-20.0,
        )

        signal = rule_consumption_drop(row)

        self.assertIsNone(signal)

    def test_drop_ignores_zero_previous_consumption(self) -> None:
        row = self.make_row(
            previous=0.0,
            latest=0.0,
            variation=-100.0,
        )

        signal = rule_consumption_drop(row)

        self.assertIsNone(signal)

    def test_drop_ignores_inconsistent_direction(self) -> None:
        row = self.make_row(
            previous=100.0,
            latest=120.0,
            variation=-40.0,
        )

        signal = rule_consumption_drop(row)

        self.assertIsNone(signal)

    def test_evaluator_returns_deterministic_order(self) -> None:
        features = pd.DataFrame(
            [
                {
                    "Material_ID": "MAT002",
                    "Material": "Material B",
                    "consumo_mes_anterior": 100.0,
                    "consumo_ultimo_mes": 180.0,
                    "variacao_consumo_pct": 80.0,
                },
                {
                    "Material_ID": "MAT001",
                    "Material": "Material A",
                    "consumo_mes_anterior": 100.0,
                    "consumo_ultimo_mes": 140.0,
                    "variacao_consumo_pct": 40.0,
                },
                {
                    "Material_ID": "MAT003",
                    "Material": "Material C",
                    "consumo_mes_anterior": 100.0,
                    "consumo_ultimo_mes": 60.0,
                    "variacao_consumo_pct": -40.0,
                },
            ]
        )

        signals = evaluate_consumption_rules(features)

        self.assertEqual(
            len(signals),
            3,
        )

        self.assertEqual(
            signals[0].id,
            "CONSUMPTION-GROWTH-MAT002",
        )

        self.assertEqual(
            signals[1].id,
            "CONSUMPTION-GROWTH-MAT001",
        )

        self.assertEqual(
            signals[2].id,
            "CONSUMPTION-DROP-MAT003",
        )

    def test_evaluator_rejects_missing_columns(self) -> None:
        features = pd.DataFrame(
            [
                {
                    "Material_ID": "MAT001",
                    "Material": "Material A",
                }
            ]
        )

        with self.assertRaises(ValueError):
            evaluate_consumption_rules(features)

    def test_invalid_numeric_value_is_ignored(self) -> None:
        row = self.make_row(
            previous=100.0,
            latest=140.0,
            variation=40.0,
        )

        row["variacao_consumo_pct"] = "invalido"

        self.assertIsNone(
            rule_consumption_growth(row)
        )

    def test_nan_value_is_ignored(self) -> None:
        row = self.make_row(
            previous=100.0,
            latest=140.0,
            variation=float("nan"),
        )

        self.assertIsNone(
            rule_consumption_growth(row)
        )


if __name__ == "__main__":
    unittest.main()