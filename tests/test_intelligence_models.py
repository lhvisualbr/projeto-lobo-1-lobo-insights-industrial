import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


from intelligence.models import (  # noqa: E402
    Evidence,
    ExecutiveReport,
    Recommendation,
    Severity,
    Signal,
)


class TestSeverity(unittest.TestCase):
    def test_ordem_de_prioridade(self):
        self.assertLess(Severity.INFO.rank, Severity.ATENCAO.rank)
        self.assertLess(Severity.ATENCAO.rank, Severity.ALTO.rank)
        self.assertLess(Severity.ALTO.rank, Severity.CRITICO.rank)

    def test_valores_publicos(self):
        self.assertEqual(Severity.INFO.value, "INFO")
        self.assertEqual(Severity.ATENCAO.value, "ATENCAO")
        self.assertEqual(Severity.ALTO.value, "ALTO")
        self.assertEqual(Severity.CRITICO.value, "CRITICO")


class TestEvidence(unittest.TestCase):
    def test_criacao_valida(self):
        evidence = Evidence(
            source="estoque",
            metric="Estoque_Atual",
            value=10,
            description="Estoque atual do material.",
        )

        self.assertEqual(evidence.source, "estoque")
        self.assertEqual(evidence.metric, "Estoque_Atual")
        self.assertEqual(evidence.value, 10)

    def test_source_vazio_falha(self):
        with self.assertRaises(ValueError):
            Evidence(
                source="",
                metric="Estoque_Atual",
                value=10,
            )

    def test_metric_vazio_falha(self):
        with self.assertRaises(ValueError):
            Evidence(
                source="estoque",
                metric="",
                value=10,
            )

    def test_to_dict(self):
        evidence = Evidence(
            source="estoque",
            metric="Estoque_Atual",
            value=10,
            description="Teste",
        )

        self.assertEqual(
            evidence.to_dict(),
            {
                "source": "estoque",
                "metric": "Estoque_Atual",
                "value": 10,
                "description": "Teste",
            },
        )


class TestRecommendation(unittest.TestCase):
    def test_criacao_valida(self):
        recommendation = Recommendation(
            action="Avaliar reposição.",
            rationale="Estoque próximo ao mínimo.",
            entity="MAT001",
            priority=Severity.ALTO,
        )

        self.assertEqual(
            recommendation.priority,
            Severity.ALTO,
        )

    def test_action_vazia_falha(self):
        with self.assertRaises(ValueError):
            Recommendation(
                action="",
                rationale="Motivo válido.",
            )

    def test_rationale_vazio_falha(self):
        with self.assertRaises(ValueError):
            Recommendation(
                action="Avaliar reposição.",
                rationale="",
            )


class TestSignal(unittest.TestCase):
    def _evidence(self):
        return Evidence(
            source="estoque",
            metric="Estoque_Atual",
            value=5,
        )

    def test_signal_valido(self):
        signal = Signal(
            id="SIG-001",
            type="ESTOQUE_BAIXO",
            severity=Severity.ALTO,
            title="Estoque baixo",
            description="Material próximo ao limite.",
            entity="MAT001",
            rule_origin="rule_stock_low",
            evidences=(self._evidence(),),
            metrics={
                "estoque_atual": 5,
                "estoque_minimo": 10,
            },
            recommended_action="Avaliar reposição.",
        )

        self.assertEqual(
            signal.priority_rank,
            Severity.ALTO.rank,
        )

    def test_signal_sem_evidencia_falha(self):
        with self.assertRaises(ValueError):
            Signal(
                id="SIG-001",
                type="ESTOQUE_BAIXO",
                severity=Severity.ALTO,
                title="Estoque baixo",
                description="Descrição.",
                entity="MAT001",
                rule_origin="rule_stock_low",
            )

    def test_campo_obrigatorio_vazio_falha(self):
        with self.assertRaises(ValueError):
            Signal(
                id="",
                type="ESTOQUE_BAIXO",
                severity=Severity.ALTO,
                title="Estoque baixo",
                description="Descrição.",
                entity="MAT001",
                rule_origin="rule_stock_low",
                evidences=(self._evidence(),),
            )

    def test_to_dict(self):
        signal = Signal(
            id="SIG-001",
            type="ESTOQUE_BAIXO",
            severity=Severity.ALTO,
            title="Estoque baixo",
            description="Material próximo ao limite.",
            entity="MAT001",
            rule_origin="rule_stock_low",
            evidences=(self._evidence(),),
            metrics={
                "estoque_atual": 5,
            },
            recommended_action="Avaliar reposição.",
        )

        data = signal.to_dict()

        self.assertEqual(data["id"], "SIG-001")
        self.assertEqual(data["severity"], "ALTO")
        self.assertEqual(
            data["evidences"][0]["metric"],
            "Estoque_Atual",
        )


class TestExecutiveReport(unittest.TestCase):
    def _signal(self, ident, severity):
        return Signal(
            id=ident,
            type="TESTE",
            severity=severity,
            title="Sinal de teste",
            description="Descrição de teste.",
            entity="MAT001",
            rule_origin="test_rule",
            evidences=(
                Evidence(
                    source="teste",
                    metric="valor",
                    value=1,
                ),
            ),
        )

    def test_total_signals(self):
        report = ExecutiveReport(
            signals=(
                self._signal("SIG-001", Severity.INFO),
                self._signal("SIG-002", Severity.CRITICO),
            )
        )

        self.assertEqual(report.total_signals, 2)

    def test_contagem_por_severidade(self):
        report = ExecutiveReport(
            signals=(
                self._signal("SIG-001", Severity.INFO),
                self._signal("SIG-002", Severity.ALTO),
                self._signal("SIG-003", Severity.ALTO),
                self._signal("SIG-004", Severity.CRITICO),
            )
        )

        self.assertEqual(
            report.count_by_severity(),
            {
                "INFO": 1,
                "ATENCAO": 0,
                "ALTO": 2,
                "CRITICO": 1,
            },
        )

    def test_ordenacao_deterministica(self):
        report = ExecutiveReport(
            signals=(
                self._signal("SIG-003", Severity.ALTO),
                self._signal("SIG-002", Severity.CRITICO),
                self._signal("SIG-001", Severity.ALTO),
            )
        )

        ordered = report.ordered_signals()

        self.assertEqual(
            [signal.id for signal in ordered],
            [
                "SIG-002",
                "SIG-001",
                "SIG-003",
            ],
        )

    def test_to_dict(self):
        report = ExecutiveReport(
            signals=(
                self._signal("SIG-001", Severity.ALTO),
            ),
            recommendations=(
                Recommendation(
                    action="Avaliar reposição.",
                    rationale="Estoque baixo.",
                    entity="MAT001",
                    priority=Severity.ALTO,
                ),
            ),
            indicators={
                "total_materiais": 10,
            },
            summary="Resumo executivo.",
        )

        data = report.to_dict()

        self.assertEqual(data["summary"], "Resumo executivo.")
        self.assertEqual(data["total_signals"], 1)
        self.assertEqual(data["indicators"]["total_materiais"], 10)
        self.assertEqual(
            data["signals"][0]["id"],
            "SIG-001",
        )


if __name__ == "__main__":
    unittest.main()