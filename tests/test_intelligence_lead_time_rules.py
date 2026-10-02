from __future__ import annotations

import sys
import unittest
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


from intelligence.lead_time_rules import (  # noqa: E402
    evaluate_lead_time_rules,
    rule_replenishment_risk,
)
from intelligence.models import Severity  # noqa: E402


class LeadTimeRulesTests(unittest.TestCase):
    def make_row(
        self,
        *,
        material_id: str = "MAT001",
        material: str = "Material Teste",
        lead_time: float = 10.0,
        coverage_months: float = 0.4,
        criticality: str = "Baixa",
    ) -> pd.Series:
        return pd.Series(
            {
                "Material_ID": material_id,
                "Material": material,
                "Lead_Time_Dias": lead_time,
                "Criticidade": criticality,
                "cobertura_meses": coverage_months,
            }
        )

    def test_factor_below_one_is_critical(
        self,
    ) -> None:
        row = self.make_row(
            lead_time=10.0,
            coverage_months=0.3,
        )

        signal = rule_replenishment_risk(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.CRITICO,
        )

        self.assertAlmostEqual(
            signal.metrics["fator_reposicao"],
            0.9,
        )

    def test_factor_equal_one_is_high(
        self,
    ) -> None:
        row = self.make_row(
            lead_time=30.0,
            coverage_months=1.0,
        )

        signal = rule_replenishment_risk(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.ALTO,
        )

    def test_factor_between_one_and_one_point_five_is_high(
        self,
    ) -> None:
        row = self.make_row(
            lead_time=10.0,
            coverage_months=0.4,
        )

        signal = rule_replenishment_risk(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.ALTO,
        )

    def test_factor_equal_one_point_five_is_attention(
        self,
    ) -> None:
        row = self.make_row(
            lead_time=20.0,
            coverage_months=1.0,
        )

        signal = rule_replenishment_risk(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.ATENCAO,
        )

    def test_factor_between_one_point_five_and_two_is_attention(
        self,
    ) -> None:
        row = self.make_row(
            lead_time=10.0,
            coverage_months=0.6,
        )

        signal = rule_replenishment_risk(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.ATENCAO,
        )

    def test_factor_equal_two_creates_no_signal(
        self,
    ) -> None:
        row = self.make_row(
            lead_time=15.0,
            coverage_months=1.0,
        )

        signal = rule_replenishment_risk(
            row
        )

        self.assertIsNone(
            signal
        )

    def test_high_criticality_escalates_attention_to_high(
        self,
    ) -> None:
        row = self.make_row(
            lead_time=10.0,
            coverage_months=0.6,
            criticality="Alta",
        )

        signal = rule_replenishment_risk(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.ALTO,
        )

        self.assertTrue(
            signal.metrics["severidade_elevada"]
        )

    def test_high_criticality_escalates_high_to_critical(
        self,
    ) -> None:
        row = self.make_row(
            lead_time=10.0,
            coverage_months=0.4,
            criticality="Alta",
        )

        signal = rule_replenishment_risk(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.CRITICO,
        )

    def test_critical_severity_does_not_exceed_critical(
        self,
    ) -> None:
        row = self.make_row(
            lead_time=10.0,
            coverage_months=0.3,
            criticality="Alta",
        )

        signal = rule_replenishment_risk(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.CRITICO,
        )

        self.assertFalse(
            signal.metrics["severidade_elevada"]
        )

    def test_signal_contains_expected_traceability(
        self,
    ) -> None:
        row = self.make_row(
            material_id="MAT099",
            material="Material Especial",
            lead_time=10.0,
            coverage_months=0.4,
        )

        signal = rule_replenishment_risk(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.id,
            "REPLENISHMENT-RISK-MAT099",
        )

        self.assertEqual(
            signal.type,
            "RISCO_REPOSICAO",
        )

        self.assertEqual(
            signal.entity,
            "MAT099",
        )

        self.assertEqual(
            signal.rule_origin,
            "rule_replenishment_risk",
        )

        self.assertEqual(
            len(signal.evidences),
            3,
        )

        self.assertAlmostEqual(
            signal.metrics["cobertura_dias"],
            12.0,
        )

    def test_zero_lead_time_is_ignored(
        self,
    ) -> None:
        row = self.make_row(
            lead_time=0.0,
        )

        self.assertIsNone(
            rule_replenishment_risk(
                row
            )
        )

    def test_negative_coverage_is_ignored(
        self,
    ) -> None:
        row = self.make_row(
            coverage_months=-1.0,
        )

        self.assertIsNone(
            rule_replenishment_risk(
                row
            )
        )

    def test_invalid_numeric_value_is_ignored(
        self,
    ) -> None:
        row = self.make_row()

        row["Lead_Time_Dias"] = "invalido"

        self.assertIsNone(
            rule_replenishment_risk(
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
                    "Lead_Time_Dias": 10.0,
                    "Criticidade": "Baixa",
                    "cobertura_meses": 0.6,
                },
                {
                    "Material_ID": "MAT002",
                    "Material": "Material B",
                    "Lead_Time_Dias": 10.0,
                    "Criticidade": "Alta",
                    "cobertura_meses": 0.4,
                },
                {
                    "Material_ID": "MAT001",
                    "Material": "Material A",
                    "Lead_Time_Dias": 10.0,
                    "Criticidade": "Baixa",
                    "cobertura_meses": 0.3,
                },
            ]
        )

        signals = evaluate_lead_time_rules(
            features
        )

        self.assertEqual(
            len(signals),
            3,
        )

        self.assertEqual(
            signals[0].id,
            "REPLENISHMENT-RISK-MAT001",
        )

        self.assertEqual(
            signals[1].id,
            "REPLENISHMENT-RISK-MAT002",
        )

        self.assertEqual(
            signals[2].id,
            "REPLENISHMENT-RISK-MAT003",
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
            evaluate_lead_time_rules(
                features
            )


if __name__ == "__main__":
    unittest.main()