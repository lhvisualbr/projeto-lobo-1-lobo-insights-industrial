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


class CostScoringTests(unittest.TestCase):
    def make_signal(
        self,
        *,
        signal_id: str = "COST-CONCENTRATION-MAT001",
        signal_type: str = "CONCENTRACAO_CUSTO",
        severity: Severity = Severity.ATENCAO,
        cost_share: object = 10.0,
        consumption_share: object = 0.0,
    ) -> Signal:
        evidences = (
            Evidence(
                source="consumo",
                metric="participacao_custo_pct",
                value=cost_share,
                description=(
                    "Participação no custo para teste."
                ),
            ),
        )

        return Signal(
            id=signal_id,
            type=signal_type,
            severity=severity,
            title="Sinal de custo para teste",
            description="Descrição de teste.",
            entity="MAT001",
            rule_origin="test_cost_scoring",
            evidences=evidences,
            metrics={
                "participacao_custo_pct": cost_share,
                "participacao_consumo_pct": consumption_share,
            },
            recommended_action=(
                "Avaliar impacto financeiro."
            ),
        )

    def test_cost_share_10_percent_scores_12_5_points(
        self,
    ) -> None:
        signal = self.make_signal(
            cost_share=10.0,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.total,
            12.5,
        )

        self.assertAlmostEqual(
            score.components["cost_share"],
            12.5,
        )

    def test_cost_share_20_percent_reaches_maximum(
        self,
    ) -> None:
        signal = self.make_signal(
            cost_share=20.0,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.total,
            25.0,
        )

        self.assertAlmostEqual(
            score.components["cost_share"],
            25.0,
        )

    def test_cost_share_above_reference_is_capped(
        self,
    ) -> None:
        signal = self.make_signal(
            cost_share=40.0,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.total,
            25.0,
        )

    def test_combined_signal_scores_cost_and_consumption(
        self,
    ) -> None:
        signal = self.make_signal(
            signal_id="COST-CONSUMPTION-IMPACT-MAT001",
            signal_type="ALTO_IMPACTO_CUSTO_CONSUMO",
            severity=Severity.ALTO,
            cost_share=12.0,
            consumption_share=14.0,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.components["cost_share"],
            15.0,
        )

        self.assertAlmostEqual(
            score.components["consumption_share"],
            14.0,
        )

        self.assertAlmostEqual(
            score.total,
            29.0,
        )

    def test_combined_components_are_capped(
        self,
    ) -> None:
        signal = self.make_signal(
            signal_id="COST-CONSUMPTION-IMPACT-MAT001",
            signal_type="ALTO_IMPACTO_CUSTO_CONSUMO",
            severity=Severity.ALTO,
            cost_share=50.0,
            consumption_share=50.0,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.components["cost_share"],
            25.0,
        )

        self.assertAlmostEqual(
            score.components["consumption_share"],
            20.0,
        )

        self.assertAlmostEqual(
            score.total,
            45.0,
        )

    def test_invalid_cost_share_scores_zero(
        self,
    ) -> None:
        signal = self.make_signal(
            cost_share="invalido",
        )

        score = score_signal(
            signal
        )

        self.assertEqual(
            score.total,
            0.0,
        )

        self.assertEqual(
            score.components["cost_share"],
            0.0,
        )

    def test_non_cost_signal_ignores_cost_metrics(
        self,
    ) -> None:
        signal = self.make_signal(
            signal_id="OTHER-MAT001",
            signal_type="OUTRO_TIPO",
            cost_share=20.0,
            consumption_share=20.0,
        )

        score = score_signal(
            signal
        )

        self.assertEqual(
            score.total,
            0.0,
        )

        self.assertNotIn(
            "cost_share",
            score.components,
        )

        self.assertNotIn(
            "consumption_share",
            score.components,
        )

    def test_cost_concentration_does_not_use_consumption_share(
        self,
    ) -> None:
        signal = self.make_signal(
            signal_type="CONCENTRACAO_CUSTO",
            cost_share=10.0,
            consumption_share=100.0,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.total,
            12.5,
        )

        self.assertNotIn(
            "consumption_share",
            score.components,
        )

    def test_same_severity_uses_cost_score_for_order(
        self,
    ) -> None:
        lower = self.make_signal(
            signal_id="COST-CONCENTRATION-MAT001",
            cost_share=10.0,
        )

        higher = self.make_signal(
            signal_id="COST-CONCENTRATION-MAT002",
            cost_share=18.0,
        )

        ranked = prioritize_signals(
            (
                lower,
                higher,
            )
        )

        self.assertEqual(
            ranked[0].signal.id,
            "COST-CONCENTRATION-MAT002",
        )

        self.assertGreater(
            ranked[0].priority.total,
            ranked[1].priority.total,
        )

    def test_severity_remains_primary_over_cost_score(
        self,
    ) -> None:
        high_severity = self.make_signal(
            signal_id="COST-CONCENTRATION-MAT001",
            severity=Severity.ALTO,
            cost_share=10.0,
        )

        high_score = self.make_signal(
            signal_id="COST-CONCENTRATION-MAT002",
            severity=Severity.ATENCAO,
            cost_share=20.0,
        )

        ranked = prioritize_signals(
            (
                high_score,
                high_severity,
            )
        )

        self.assertEqual(
            ranked[0].signal.id,
            "COST-CONCENTRATION-MAT001",
        )

        self.assertEqual(
            ranked[0].signal.severity,
            Severity.ALTO,
        )

        self.assertLess(
            ranked[0].priority.total,
            ranked[1].priority.total,
        )


if __name__ == "__main__":
    unittest.main()