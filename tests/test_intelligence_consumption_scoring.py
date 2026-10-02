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


class ConsumptionScoringTests(unittest.TestCase):
    def make_signal(
        self,
        *,
        signal_id: str = "CONSUMPTION-GROWTH-MAT001",
        signal_type: str = "CRESCIMENTO_CONSUMO",
        severity: Severity = Severity.ATENCAO,
        variation: object = 40.0,
    ) -> Signal:
        evidence = Evidence(
            source="consumo",
            metric="variacao_consumo_pct",
            value=variation,
            description="Variação de consumo para teste.",
        )

        return Signal(
            id=signal_id,
            type=signal_type,
            severity=severity,
            title="Sinal de consumo para teste",
            description="Descrição de teste.",
            entity="MAT001",
            rule_origin="test_consumption_scoring",
            evidences=(evidence,),
            metrics={
                "variacao_consumo_pct": variation,
            },
            recommended_action=(
                "Avaliar comportamento de consumo."
            ),
        )

    def test_growth_40_percent_scores_10_points(
        self,
    ) -> None:
        signal = self.make_signal(
            variation=40.0,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.total,
            10.0,
        )

        self.assertAlmostEqual(
            score.components[
                "consumption_variation"
            ],
            10.0,
        )

    def test_drop_50_percent_scores_12_5_points(
        self,
    ) -> None:
        signal = self.make_signal(
            signal_id="CONSUMPTION-DROP-MAT001",
            signal_type="REDUCAO_CONSUMO",
            variation=-50.0,
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
                "consumption_variation"
            ],
            12.5,
        )

    def test_absolute_magnitude_is_used_for_drop(
        self,
    ) -> None:
        growth = self.make_signal(
            variation=60.0,
        )

        drop = self.make_signal(
            signal_id="CONSUMPTION-DROP-MAT001",
            signal_type="REDUCAO_CONSUMO",
            variation=-60.0,
        )

        growth_score = score_signal(
            growth
        )

        drop_score = score_signal(
            drop
        )

        self.assertAlmostEqual(
            growth_score.total,
            drop_score.total,
        )

    def test_100_percent_reaches_component_maximum(
        self,
    ) -> None:
        signal = self.make_signal(
            variation=100.0,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.total,
            25.0,
        )

        self.assertAlmostEqual(
            score.components[
                "consumption_variation"
            ],
            25.0,
        )

    def test_variation_above_100_percent_is_capped(
        self,
    ) -> None:
        signal = self.make_signal(
            variation=250.0,
        )

        score = score_signal(
            signal
        )

        self.assertAlmostEqual(
            score.total,
            25.0,
        )

        self.assertAlmostEqual(
            score.components[
                "consumption_variation"
            ],
            25.0,
        )

    def test_invalid_variation_scores_zero(
        self,
    ) -> None:
        signal = self.make_signal(
            variation="invalido",
        )

        score = score_signal(
            signal
        )

        self.assertEqual(
            score.total,
            0.0,
        )

        self.assertEqual(
            score.components[
                "consumption_variation"
            ],
            0.0,
        )

    def test_non_consumption_signal_ignores_consumption_metric(
        self,
    ) -> None:
        signal = self.make_signal(
            signal_id="OTHER-MAT001",
            signal_type="OUTRO_TIPO",
            variation=100.0,
        )

        score = score_signal(
            signal
        )

        self.assertEqual(
            score.total,
            0.0,
        )

        self.assertNotIn(
            "consumption_variation",
            score.components,
        )

    def test_same_severity_uses_consumption_score_for_order(
        self,
    ) -> None:
        lower = self.make_signal(
            signal_id="CONSUMPTION-GROWTH-MAT001",
            variation=40.0,
        )

        higher = self.make_signal(
            signal_id="CONSUMPTION-GROWTH-MAT002",
            variation=80.0,
        )

        ranked = prioritize_signals(
            (
                lower,
                higher,
            )
        )

        self.assertEqual(
            ranked[0].signal.id,
            "CONSUMPTION-GROWTH-MAT002",
        )

        self.assertGreater(
            ranked[0].priority.total,
            ranked[1].priority.total,
        )

    def test_severity_remains_primary_over_score(
        self,
    ) -> None:
        high_severity = self.make_signal(
            signal_id="CONSUMPTION-GROWTH-MAT001",
            severity=Severity.ALTO,
            variation=30.0,
        )

        high_score = self.make_signal(
            signal_id="CONSUMPTION-GROWTH-MAT002",
            severity=Severity.ATENCAO,
            variation=100.0,
        )

        ranked = prioritize_signals(
            (
                high_score,
                high_severity,
            )
        )

        self.assertEqual(
            ranked[0].signal.id,
            "CONSUMPTION-GROWTH-MAT001",
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