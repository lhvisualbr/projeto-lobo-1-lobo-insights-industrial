import sys
import unittest
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
SCRIPTS = ROOT / "scripts"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


from intelligence.engine import (  # noqa: E402
    EngineResult,
    run_engine,
)
from intelligence.models import (  # noqa: E402
    ExecutiveReport,
)
from lobo_common import DATA, ENC, SEP  # noqa: E402


class TestExecutiveIntelligenceEngine(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.estoque = pd.read_csv(
            DATA / "estoque_ficticio.csv",
            sep=SEP,
            encoding=ENC,
        )

        cls.consumo = pd.read_csv(
            DATA / "consumo_ficticio.csv",
            sep=SEP,
            encoding=ENC,
        )

    def run_real_engine(self):
        return run_engine(
            self.estoque,
            self.consumo,
        )

    def test_execucao_retorna_engine_result(self):
        result = self.run_real_engine()

        self.assertIsInstance(
            result,
            EngineResult,
        )

        self.assertIsInstance(
            result.report,
            ExecutiveReport,
        )

    def test_contexto_real_da_base(self):
        result = self.run_real_engine()

        self.assertEqual(
            result.context.data_inicio,
            "2026-06-01",
        )

        self.assertEqual(
            result.context.data_fim,
            "2026-08-31",
        )

        self.assertEqual(
            result.context.total_movimentacoes,
            180,
        )

        self.assertAlmostEqual(
            result.context.total_quantidade,
            1647.0,
        )

        self.assertAlmostEqual(
            result.context.total_custo,
            24606.20,
            places=2,
        )

    def test_indicadores_basicos_sao_coerentes(self):
        result = self.run_real_engine()

        indicators = result.report.indicators

        self.assertEqual(
            indicators["materiais_analisados"],
            20,
        )

        self.assertEqual(
            indicators["movimentacoes_analisadas"],
            180,
        )

        self.assertAlmostEqual(
            indicators["quantidade_total_consumida"],
            1647.0,
        )

        self.assertAlmostEqual(
            indicators["custo_total_consumo"],
            24606.20,
            places=2,
        )

    def test_total_de_sinais_bate_com_relatorio(self):
        result = self.run_real_engine()

        self.assertEqual(
            result.report.total_signals,
            len(result.ranked_signals),
        )

        self.assertEqual(
            result.report.indicators["total_sinais"],
            result.report.total_signals,
        )

    def test_base_real_produz_sinais(self):
        result = self.run_real_engine()

        self.assertGreater(
            result.report.total_signals,
            0,
        )

        entities = {
            signal.entity
            for signal in result.report.signals
        }

        self.assertIn(
            "MAT005",
            entities,
        )

        self.assertIn(
            "MAT007",
            entities,
        )

    def test_ordem_do_relatorio_preserva_ranking(self):
        result = self.run_real_engine()

        ids_report = [
            signal.id
            for signal in result.report.signals
        ]

        ids_ranking = [
            item.signal.id
            for item in result.ranked_signals
        ]

        self.assertEqual(
            ids_report,
            ids_ranking,
        )

    def test_priority_score_e_incorporado_ao_relatorio(self):
        result = self.run_real_engine()

        for signal, ranked in zip(
            result.report.signals,
            result.ranked_signals,
        ):
            self.assertIn(
                "priority_score",
                signal.metrics,
            )

            self.assertEqual(
                signal.metrics["priority_score"],
                ranked.priority.total,
            )

    def test_recomendacoes_nao_possuem_duplicacao_exata(self):
        result = self.run_real_engine()

        keys = [
            (
                recommendation.entity,
                recommendation.action,
            )
            for recommendation
            in result.report.recommendations
        ]

        self.assertEqual(
            len(keys),
            len(set(keys)),
        )

    def test_recomendacoes_sao_rastreaveis_a_sinais(self):
        result = self.run_real_engine()

        expected = {
            (
                signal.entity,
                signal.recommended_action,
            )
            for signal in result.report.signals
            if signal.recommended_action
        }

        actual = {
            (
                recommendation.entity,
                recommendation.action,
            )
            for recommendation
            in result.report.recommendations
        }

        self.assertEqual(
            actual,
            expected,
        )

    def test_execucao_e_deterministica(self):
        first = self.run_real_engine()
        second = self.run_real_engine()

        self.assertEqual(
            first.to_dict(),
            second.to_dict(),
        )

    def test_dataframes_de_entrada_nao_sao_modificados(self):
        estoque_before = self.estoque.copy(
            deep=True
        )

        consumo_before = self.consumo.copy(
            deep=True
        )

        self.run_real_engine()

        pd.testing.assert_frame_equal(
            self.estoque,
            estoque_before,
        )

        pd.testing.assert_frame_equal(
            self.consumo,
            consumo_before,
        )

    def test_serializacao_possui_estrutura_executiva(self):
        result = self.run_real_engine()

        data = result.to_dict()

        self.assertIn(
            "report",
            data,
        )

        self.assertIn(
            "ranking",
            data,
        )

        self.assertIn(
            "context",
            data,
        )

        self.assertIn(
            "summary",
            data["report"],
        )

        self.assertIn(
            "signals",
            data["report"],
        )

        self.assertIn(
            "recommendations",
            data["report"],
        )

        self.assertIn(
            "indicators",
            data["report"],
        )

    def test_resumo_menciona_maior_prioridade(self):
        result = self.run_real_engine()

        self.assertTrue(
            result.ranked_signals
        )

        top = result.ranked_signals[0]

        self.assertIn(
            top.signal.entity,
            result.report.summary,
        )

        self.assertIn(
            top.signal.severity.value,
            result.report.summary,
        )


if __name__ == "__main__":
    unittest.main()