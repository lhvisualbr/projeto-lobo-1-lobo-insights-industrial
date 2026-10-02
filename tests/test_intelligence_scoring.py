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


def make_signal(
    *,
    ident="SIG-001",
    severity=Severity.ALTO,
    criticidade="Baixa",
    ratio=2.0,
    cobertura=2.0,
    lead_time=0.0,
):
    return Signal(
        id=ident,
        type="TESTE",
        severity=severity,
        title="Sinal de teste",
        description="Sinal criado para teste de scoring.",
        entity="MAT001",
        rule_origin="test_scoring",
        evidences=(
            Evidence(
                source="estoque",
                metric="Criticidade",
                value=criticidade,
            ),
        ),
        metrics={
            "razao_estoque_minimo": ratio,
            "cobertura_meses": cobertura,
            "lead_time_dias": lead_time,
        },
    )


class TestScoreSignal(unittest.TestCase):
    def test_score_sem_fatores_de_pressao_e_zero(self):
        signal = make_signal()

        result = score_signal(signal)

        self.assertEqual(
            result.total,
            0.0,
        )

    def test_criticidade_alta_adiciona_25(self):
        signal = make_signal(
            criticidade="Alta",
        )

        result = score_signal(signal)

        self.assertEqual(
            result.components["criticality"],
            25.0,
        )

        self.assertEqual(
            result.total,
            25.0,
        )

    def test_criticidade_media_adiciona_12_5(self):
        signal = make_signal(
            criticidade="Média",
        )

        result = score_signal(signal)

        self.assertEqual(
            result.components["criticality"],
            12.5,
        )

    def test_pressao_estoque_abaixo_minimo(self):
        signal = make_signal(
            ratio=0.5,
        )

        result = score_signal(signal)

        self.assertEqual(
            result.components["stock_pressure"],
            22.5,
        )

    def test_pressao_estoque_proximo_minimo(self):
        signal = make_signal(
            ratio=1.10,
        )

        result = score_signal(signal)

        self.assertEqual(
            result.components["stock_pressure"],
            9.0,
        )

    def test_baixa_cobertura(self):
        signal = make_signal(
            cobertura=0.4,
        )

        result = score_signal(signal)

        self.assertEqual(
            result.components["coverage_pressure"],
            15.0,
        )

    def test_lead_time_15_dias(self):
        signal = make_signal(
            lead_time=15,
        )

        result = score_signal(signal)

        self.assertEqual(
            result.components["lead_time"],
            10.0,
        )

    def test_score_fica_entre_zero_e_cem(self):
        signal = make_signal(
            criticidade="Alta",
            ratio=-10,
            cobertura=-10,
            lead_time=100,
        )

        result = score_signal(signal)

        self.assertGreaterEqual(
            result.total,
            0.0,
        )

        self.assertLessEqual(
            result.total,
            100.0,
        )

        self.assertEqual(
            result.total,
            100.0,
        )

    def test_scoring_e_deterministico(self):
        signal = make_signal(
            criticidade="Alta",
            ratio=0.7,
            cobertura=0.6,
            lead_time=22,
        )

        first = score_signal(signal)
        second = score_signal(signal)

        self.assertEqual(
            first,
            second,
        )

    def test_componentes_geram_explicacoes(self):
        signal = make_signal(
            criticidade="Alta",
            ratio=0.5,
            cobertura=0.5,
            lead_time=30,
        )

        result = score_signal(signal)

        self.assertEqual(
            len(result.explanations),
            4,
        )

        self.assertTrue(
            all(
                explanation.strip()
                for explanation in result.explanations
            )
        )

    def test_priority_score_to_dict(self):
        signal = make_signal(
            criticidade="Alta",
        )

        data = score_signal(
            signal
        ).to_dict()

        self.assertEqual(
            data["total"],
            25.0,
        )

        self.assertIn(
            "criticality",
            data["components"],
        )

        self.assertIsInstance(
            data["explanations"],
            list,
        )


class TestPrioritizeSignals(unittest.TestCase):
    def test_severidade_tem_prioridade_sobre_score(self):
        critical = make_signal(
            ident="SIG-CRITICO",
            severity=Severity.CRITICO,
            criticidade="Baixa",
            ratio=2.0,
            cobertura=2.0,
            lead_time=0,
        )

        high_score = make_signal(
            ident="SIG-ALTO",
            severity=Severity.ALTO,
            criticidade="Alta",
            ratio=-10,
            cobertura=-10,
            lead_time=100,
        )

        ranked = prioritize_signals(
            [
                high_score,
                critical,
            ]
        )

        self.assertEqual(
            ranked[0].signal.id,
            "SIG-CRITICO",
        )

        self.assertEqual(
            ranked[1].signal.id,
            "SIG-ALTO",
        )

    def test_score_desempata_mesma_severidade(self):
        lower = make_signal(
            ident="SIG-A",
            severity=Severity.ALTO,
        )

        higher = make_signal(
            ident="SIG-Z",
            severity=Severity.ALTO,
            criticidade="Alta",
            ratio=0.5,
            cobertura=0.5,
            lead_time=30,
        )

        ranked = prioritize_signals(
            [
                lower,
                higher,
            ]
        )

        self.assertEqual(
            ranked[0].signal.id,
            "SIG-Z",
        )

        self.assertGreater(
            ranked[0].priority.total,
            ranked[1].priority.total,
        )

    def test_id_desempata_scores_identicos(self):
        signal_b = make_signal(
            ident="SIG-B",
            severity=Severity.ALTO,
        )

        signal_a = make_signal(
            ident="SIG-A",
            severity=Severity.ALTO,
        )

        ranked = prioritize_signals(
            [
                signal_b,
                signal_a,
            ]
        )

        self.assertEqual(
            [
                item.signal.id
                for item in ranked
            ],
            [
                "SIG-A",
                "SIG-B",
            ],
        )

    def test_ranked_signal_to_dict(self):
        signal = make_signal(
            ident="SIG-001",
            criticidade="Alta",
        )

        ranked = prioritize_signals(
            [signal]
        )[0]

        data = ranked.to_dict()

        self.assertEqual(
            data["signal"]["id"],
            "SIG-001",
        )

        self.assertEqual(
            data["priority"]["total"],
            25.0,
        )


if __name__ == "__main__":
    unittest.main()