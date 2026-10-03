from __future__ import annotations

import sys
import unittest
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


from intelligence.models import Severity  # noqa: E402
from intelligence.supplier_rules import (  # noqa: E402
    build_supplier_features,
    evaluate_supplier_rules,
    rule_supplier_cost_concentration,
    rule_supplier_critical_exposure,
)


class SupplierRulesTests(unittest.TestCase):
    def make_supplier_row(
        self,
        *,
        supplier: str = "Fornecedor Teste",
        materials: int = 4,
        critical_materials: int = 2,
        total_cost: float = 5000.0,
        cost_share: float = 30.0,
    ) -> pd.Series:
        return pd.Series(
            {
                "Fornecedor": supplier,
                "materiais": materials,
                "materiais_criticos": critical_materials,
                "custo_total": total_cost,
                "participacao_custo_pct": cost_share,
            }
        )

    def test_build_supplier_features_aggregates_materials(
        self,
    ) -> None:
        features = pd.DataFrame(
            [
                {
                    "Material_ID": "MAT001",
                    "Fornecedor": "Fornecedor Alfa",
                    "Criticidade": "Alta",
                    "custo_consumo": 1000.0,
                    "participacao_custo_pct": 10.0,
                },
                {
                    "Material_ID": "MAT002",
                    "Fornecedor": "Fornecedor Alfa",
                    "Criticidade": "Média",
                    "custo_consumo": 2000.0,
                    "participacao_custo_pct": 20.0,
                },
                {
                    "Material_ID": "MAT003",
                    "Fornecedor": "Fornecedor Beta",
                    "Criticidade": "Alta",
                    "custo_consumo": 500.0,
                    "participacao_custo_pct": 5.0,
                },
            ]
        )

        result = build_supplier_features(
            features
        )

        self.assertEqual(
            len(result),
            2,
        )

        alfa = result[
            result["Fornecedor"]
            == "Fornecedor Alfa"
        ].iloc[0]

        self.assertEqual(
            alfa["materiais"],
            2,
        )

        self.assertEqual(
            alfa["materiais_criticos"],
            1,
        )

        self.assertAlmostEqual(
            alfa["custo_total"],
            3000.0,
        )

        self.assertAlmostEqual(
            alfa["participacao_custo_pct"],
            30.0,
        )

    def test_build_supplier_features_ignores_blank_supplier(
        self,
    ) -> None:
        features = pd.DataFrame(
            [
                {
                    "Material_ID": "MAT001",
                    "Fornecedor": "",
                    "Criticidade": "Alta",
                    "custo_consumo": 1000.0,
                    "participacao_custo_pct": 10.0,
                },
                {
                    "Material_ID": "MAT002",
                    "Fornecedor": "Fornecedor Alfa",
                    "Criticidade": "Média",
                    "custo_consumo": 2000.0,
                    "participacao_custo_pct": 20.0,
                },
            ]
        )

        result = build_supplier_features(
            features
        )

        self.assertEqual(
            len(result),
            1,
        )

        self.assertEqual(
            result.iloc[0]["Fornecedor"],
            "Fornecedor Alfa",
        )

    def test_cost_concentration_creates_attention_signal(
        self,
    ) -> None:
        row = self.make_supplier_row(
            cost_share=30.0,
        )

        signal = rule_supplier_cost_concentration(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.type,
            "CONCENTRACAO_FORNECEDOR",
        )

        self.assertEqual(
            signal.severity,
            Severity.ATENCAO,
        )

        self.assertEqual(
            signal.entity,
            "Fornecedor Teste",
        )

        self.assertEqual(
            signal.rule_origin,
            "rule_supplier_cost_concentration",
        )

        self.assertEqual(
            len(signal.evidences),
            3,
        )

    def test_cost_concentration_becomes_high_at_40_percent(
        self,
    ) -> None:
        row = self.make_supplier_row(
            cost_share=40.0,
        )

        signal = rule_supplier_cost_concentration(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.ALTO,
        )

    def test_cost_concentration_accepts_exact_25_percent(
        self,
    ) -> None:
        row = self.make_supplier_row(
            cost_share=25.0,
        )

        signal = rule_supplier_cost_concentration(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.ATENCAO,
        )

    def test_cost_concentration_ignores_below_threshold(
        self,
    ) -> None:
        row = self.make_supplier_row(
            cost_share=24.99,
        )

        self.assertIsNone(
            rule_supplier_cost_concentration(
                row
            )
        )

    def test_critical_exposure_creates_high_signal(
        self,
    ) -> None:
        row = self.make_supplier_row(
            critical_materials=2,
        )

        signal = rule_supplier_critical_exposure(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.type,
            "EXPOSICAO_FORNECEDOR_CRITICO",
        )

        self.assertEqual(
            signal.severity,
            Severity.ALTO,
        )

        self.assertEqual(
            signal.rule_origin,
            "rule_supplier_critical_exposure",
        )

        self.assertEqual(
            len(signal.evidences),
            3,
        )

    def test_critical_exposure_becomes_critical_at_three(
        self,
    ) -> None:
        row = self.make_supplier_row(
            critical_materials=3,
        )

        signal = rule_supplier_critical_exposure(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.severity,
            Severity.CRITICO,
        )

    def test_critical_exposure_ignores_one_material(
        self,
    ) -> None:
        row = self.make_supplier_row(
            critical_materials=1,
        )

        self.assertIsNone(
            rule_supplier_critical_exposure(
                row
            )
        )

    def test_supplier_id_is_deterministic(
        self,
    ) -> None:
        row = self.make_supplier_row(
            supplier="Fornecedor Álfä & Cia",
            cost_share=45.0,
        )

        signal = rule_supplier_cost_concentration(
            row
        )

        self.assertIsNotNone(signal)
        assert signal is not None

        self.assertEqual(
            signal.id,
            "SUPPLIER-COST-CONCENTRATION-FORNECEDOR-ALFA-CIA",
        )

    def test_invalid_cost_share_is_ignored(
        self,
    ) -> None:
        row = self.make_supplier_row()

        row[
            "participacao_custo_pct"
        ] = "invalido"

        self.assertIsNone(
            rule_supplier_cost_concentration(
                row
            )
        )

    def test_zero_total_cost_is_ignored(
        self,
    ) -> None:
        row = self.make_supplier_row(
            total_cost=0.0,
            cost_share=50.0,
        )

        self.assertIsNone(
            rule_supplier_cost_concentration(
                row
            )
        )

    def test_evaluator_returns_expected_signals_and_order(
        self,
    ) -> None:
        features = pd.DataFrame(
            [
                {
                    "Material_ID": "MAT001",
                    "Fornecedor": "Fornecedor Alfa",
                    "Criticidade": "Alta",
                    "custo_consumo": 2500.0,
                    "participacao_custo_pct": 25.0,
                },
                {
                    "Material_ID": "MAT002",
                    "Fornecedor": "Fornecedor Alfa",
                    "Criticidade": "Alta",
                    "custo_consumo": 2000.0,
                    "participacao_custo_pct": 20.0,
                },
                {
                    "Material_ID": "MAT003",
                    "Fornecedor": "Fornecedor Beta",
                    "Criticidade": "Baixa",
                    "custo_consumo": 3000.0,
                    "participacao_custo_pct": 30.0,
                },
                {
                    "Material_ID": "MAT004",
                    "Fornecedor": "Fornecedor Beta",
                    "Criticidade": "Média",
                    "custo_consumo": 500.0,
                    "participacao_custo_pct": 5.0,
                },
            ]
        )

        signals = evaluate_supplier_rules(
            features
        )

        self.assertEqual(
            len(signals),
            3,
        )

        self.assertEqual(
            signals[0].id,
            "SUPPLIER-COST-CONCENTRATION-FORNECEDOR-ALFA",
        )

        self.assertEqual(
            signals[1].id,
            "SUPPLIER-CRITICAL-EXPOSURE-FORNECEDOR-ALFA",
        )

        self.assertEqual(
            signals[2].id,
            "SUPPLIER-COST-CONCENTRATION-FORNECEDOR-BETA",
        )

        self.assertEqual(
            signals[0].severity,
            Severity.ALTO,
        )

        self.assertEqual(
            signals[1].severity,
            Severity.ALTO,
        )

        self.assertEqual(
            signals[2].severity,
            Severity.ATENCAO,
        )

    def test_evaluator_rejects_missing_columns(
        self,
    ) -> None:
        features = pd.DataFrame(
            [
                {
                    "Material_ID": "MAT001",
                    "Fornecedor": "Fornecedor Alfa",
                }
            ]
        )

        with self.assertRaises(
            ValueError
        ):
            evaluate_supplier_rules(
                features
            )


if __name__ == "__main__":
    unittest.main()