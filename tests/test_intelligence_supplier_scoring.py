from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


from intelligence.models import (  # noqa: E402
    Evidence,
    Severity,
    Signal,
)
from intelligence.scoring import (  # noqa: E402
    prioritize_signals,
    score_signal,
)


class SupplierScoringTests(unittest.TestCase):
    def make_signal(
        self,
        *,
        signal_id: str = (
            "SUPPLIER-COST-CONCENTRATION-FORNECEDOR-ALFA"
        ),
        signal_type: str = "CONCENTRACAO_FORNECEDOR",
        severity: Severity = Severity.ATENCAO,
        cost_share: object = 25.0,
        critical_materials: object = 0,
    ) -> Signal:
        evidences = (
            Evidence(
                source="fornecedor",
                metric="participacao_custo_pct",
                value=cost_share,
                description="Participação financeira para teste.",
            ),
        )

        return Signal(
            id=signal_id,
            type=signal_type,
            severity=severity,
            title="Sinal de fornecedor para teste",
            description="Descrição de teste.",
            entity="Fornecedor Alfa",
            rule_origin="test_supplier_scoring",
            evidences=evidences,
            metrics={
                "participacao_custo_pct": cost_share,
                "materiais_criticos": critical_materials,
            },
            recommended_action=(
                "Avaliar dependência do fornecedor."
            ),
        )

    def test_supplier_25_percent_scores_12_5_points(
        self,
    ) -> None:
        signal = self.make_signal(
            cost_share=25.0,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.total,
            12.5,
        )

        self.assertAlmostEqual(
            score.components[
                "supplier_cost_share"
            ],
            12.5,
        )

    def test_supplier_50_percent_reaches_cost_maximum(
        self,
    ) -> None:
        signal = self.make_signal(
            cost_share=50.0,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.total,
            25.0,
        )

    def test_supplier_cost_above_reference_is_capped(
        self,
    ) -> None:
        signal = self.make_signal(
            cost_share=80.0,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.total,
            25.0,
        )

    def test_two_critical_materials_score_20_points(
        self,
    ) -> None:
        signal = self.make_signal(
            signal_id=(
                "SUPPLIER-CRITICAL-EXPOSURE-FORNECEDOR-ALFA"
            ),
            signal_type="EXPOSICAO_FORNECEDOR_CRITICO",
            severity=Severity.ALTO,
            cost_share=0.0,
            critical_materials=2,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.components[
                "supplier_critical_exposure"
            ],
            20.0,
        )

        self.assertAlmostEqual(
            score.total,
            20.0,
        )

    def test_three_critical_materials_reach_maximum(
        self,
    ) -> None:
        signal = self.make_signal(
            signal_id=(
                "SUPPLIER-CRITICAL-EXPOSURE-FORNECEDOR-ALFA"
            ),
            signal_type="EXPOSICAO_FORNECEDOR_CRITICO",
            severity=Severity.CRITICO,
            cost_share=0.0,
            critical_materials=3,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.components[
                "supplier_critical_exposure"
            ],
            30.0,
        )

    def test_critical_exposure_combines_cost_and_criticality(
        self,
    ) -> None:
        signal = self.make_signal(
            signal_id=(
                "SUPPLIER-CRITICAL-EXPOSURE-FORNECEDOR-ALFA"
            ),
            signal_type="EXPOSICAO_FORNECEDOR_CRITICO",
            severity=Severity.ALTO,
            cost_share=40.0,
            critical_materials=2,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.components[
                "supplier_cost_share"
            ],
            20.0,
        )

        self.assertAlmostEqual(
            score.components[
                "supplier_critical_exposure"
            ],
            20.0,
        )

        self.assertAlmostEqual(
            score.total,
            40.0,
        )

    def test_non_supplier_signal_ignores_supplier_metrics(
        self,
    ) -> None:
        signal = self.make_signal(
            signal_id="OTHER-001",
            signal_type="OUTRO_TIPO",
            cost_share=50.0,
            critical_materials=3,
        )

        score = score_signal(
            signal
        )

        self.assertEqual(
            score.total,
            0.0,
        )

        self.assertNotIn(
            "supplier_cost_share",
            score.components,
        )

        self.assertNotIn(
            "supplier_critical_exposure",
            score.components,
        )

    def test_same_severity_uses_supplier_score_for_order(
        self,
    ) -> None:
        lower = self.make_signal(
            signal_id=(
                "SUPPLIER-COST-CONCENTRATION-BETA"
            ),
            severity=Severity.ALTO,
            cost_share=40.0,
        )

        higher = self.make_signal(
            signal_id=(
                "SUPPLIER-CRITICAL-EXPOSURE-ALFA"
            ),
            signal_type="EXPOSICAO_FORNECEDOR_CRITICO",
            severity=Severity.ALTO,
            cost_share=40.0,
            critical_materials=2,
        )

        ranked = prioritize_signals(
            (
                lower,
                higher,
            )
        )

        self.assertEqual(
            ranked[0].signal.id,
            "SUPPLIER-CRITICAL-EXPOSURE-ALFA",
        )

    def test_severity_remains_primary_over_supplier_score(
        self,
    ) -> None:
        critical = self.make_signal(
            signal_id="SUPPLIER-CRITICAL",
            severity=Severity.CRITICO,
            cost_share=25.0,
        )

        high_score = self.make_signal(
            signal_id="SUPPLIER-HIGH",
            signal_type="EXPOSICAO_FORNECEDOR_CRITICO",
            severity=Severity.ALTO,
            cost_share=50.0,
            critical_materials=3,
        )

        ranked = prioritize_signals(
            (
                high_score,
                critical,
            )
        )

        self.assertEqual(
            ranked[0].signal.id,
            "SUPPLIER-CRITICAL",
        )


if __name__ == "__main__":
    unittest.main()