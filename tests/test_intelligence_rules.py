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


from intelligence.features import build_material_features  # noqa: E402
from intelligence.models import Severity  # noqa: E402
from intelligence.rules import (  # noqa: E402
    evaluate_inventory_rules,
    rule_low_coverage,
    rule_stock_below_minimum,
    rule_stock_near_minimum,
)
from lobo_common import DATA, ENC, SEP  # noqa: E402


def make_row(**overrides):
    row = {
        "Material_ID": "MAT-TESTE",
        "Material": "Material de Teste",
        "Criticidade": "Média",
        "Estoque_Atual": 20,
        "Estoque_Minimo": 10,
        "Lead_Time_Dias": 15,
        "gap_estoque_minimo": 10,
        "razao_estoque_minimo": 2.0,
        "cobertura_meses": 2.0,
        "media_mensal_consumo": 10.0,
    }

    row.update(overrides)
    return row


class TestStockBelowMinimum(unittest.TestCase):
    def test_criticidade_alta_gera_critico(self):
        row = make_row(
            Criticidade="Alta",
            Estoque_Atual=5,
            Estoque_Minimo=10,
        )

        signal = rule_stock_below_minimum(row)

        self.assertIsNotNone(signal)
        self.assertEqual(
            signal.severity,
            Severity.CRITICO,
        )
        self.assertEqual(
            signal.type,
            "ESTOQUE_ABAIXO_MINIMO",
        )

    def test_criticidade_nao_alta_gera_alto(self):
        row = make_row(
            Criticidade="Média",
            Estoque_Atual=5,
            Estoque_Minimo=10,
        )

        signal = rule_stock_below_minimum(row)

        self.assertIsNotNone(signal)
        self.assertEqual(
            signal.severity,
            Severity.ALTO,
        )

    def test_estoque_igual_ao_minimo_nao_gera_sinal(self):
        row = make_row(
            Estoque_Atual=10,
            Estoque_Minimo=10,
        )

        self.assertIsNone(
            rule_stock_below_minimum(row)
        )

    def test_sinal_possui_evidencias_objetivas(self):
        row = make_row(
            Estoque_Atual=5,
            Estoque_Minimo=10,
        )

        signal = rule_stock_below_minimum(row)

        metrics = {
            evidence.metric
            for evidence in signal.evidences
        }

        self.assertIn(
            "Estoque_Atual",
            metrics,
        )

        self.assertIn(
            "Estoque_Minimo",
            metrics,
        )

        self.assertIn(
            "Criticidade",
            metrics,
        )


class TestStockNearMinimum(unittest.TestCase):
    def test_criticidade_alta_eleva_para_alto(self):
        row = make_row(
            Criticidade="Alta",
            Estoque_Atual=12,
            Estoque_Minimo=10,
            razao_estoque_minimo=1.2,
        )

        signal = rule_stock_near_minimum(row)

        self.assertIsNotNone(signal)
        self.assertEqual(
            signal.severity,
            Severity.ALTO,
        )

    def test_criticidade_normal_gera_atencao(self):
        row = make_row(
            Criticidade="Média",
            Estoque_Atual=12,
            Estoque_Minimo=10,
            razao_estoque_minimo=1.2,
        )

        signal = rule_stock_near_minimum(row)

        self.assertIsNotNone(signal)
        self.assertEqual(
            signal.severity,
            Severity.ATENCAO,
        )

    def test_abaixo_do_minimo_nao_duplica_regra(self):
        row = make_row(
            Estoque_Atual=8,
            Estoque_Minimo=10,
            razao_estoque_minimo=0.8,
        )

        self.assertIsNone(
            rule_stock_near_minimum(row)
        )

    def test_acima_da_faixa_nao_gera_sinal(self):
        row = make_row(
            Estoque_Atual=13,
            Estoque_Minimo=10,
            razao_estoque_minimo=1.3,
        )

        self.assertIsNone(
            rule_stock_near_minimum(row)
        )


class TestLowCoverage(unittest.TestCase):
    def test_baixa_cobertura_criticidade_alta_gera_critico(self):
        row = make_row(
            Criticidade="Alta",
            cobertura_meses=0.8,
            media_mensal_consumo=25,
        )

        signal = rule_low_coverage(row)

        self.assertIsNotNone(signal)
        self.assertEqual(
            signal.severity,
            Severity.CRITICO,
        )

    def test_baixa_cobertura_normal_gera_alto(self):
        row = make_row(
            Criticidade="Baixa",
            cobertura_meses=0.8,
            media_mensal_consumo=25,
        )

        signal = rule_low_coverage(row)

        self.assertIsNotNone(signal)
        self.assertEqual(
            signal.severity,
            Severity.ALTO,
        )

    def test_cobertura_no_limite_nao_gera_sinal(self):
        row = make_row(
            cobertura_meses=1.0,
            media_mensal_consumo=20,
        )

        self.assertIsNone(
            rule_low_coverage(row)
        )

    def test_consumo_zero_nao_inventa_risco(self):
        row = make_row(
            cobertura_meses=0.5,
            media_mensal_consumo=0,
        )

        self.assertIsNone(
            rule_low_coverage(row)
        )

    def test_lead_time_entra_como_evidencia(self):
        row = make_row(
            cobertura_meses=0.7,
            media_mensal_consumo=20,
            Lead_Time_Dias=30,
        )

        signal = rule_low_coverage(row)

        metrics = {
            evidence.metric
            for evidence in signal.evidences
        }

        self.assertIn(
            "Lead_Time_Dias",
            metrics,
        )


class TestEvaluateInventoryRules(unittest.TestCase):
    def test_coluna_obrigatoria_ausente_falha(self):
        df = pd.DataFrame(
            [
                make_row()
            ]
        ).drop(
            columns=["Lead_Time_Dias"]
        )

        with self.assertRaises(ValueError):
            evaluate_inventory_rules(df)

    def test_ordenacao_e_deterministica(self):
        df = pd.DataFrame(
            [
                make_row(
                    Material_ID="MAT002",
                    Material="Material B",
                    Criticidade="Média",
                    Estoque_Atual=5,
                    Estoque_Minimo=10,
                    gap_estoque_minimo=-5,
                    razao_estoque_minimo=0.5,
                    cobertura_meses=0.5,
                ),
                make_row(
                    Material_ID="MAT001",
                    Material="Material A",
                    Criticidade="Alta",
                    Estoque_Atual=5,
                    Estoque_Minimo=10,
                    gap_estoque_minimo=-5,
                    razao_estoque_minimo=0.5,
                    cobertura_meses=0.5,
                ),
            ]
        )

        first = evaluate_inventory_rules(df)
        second = evaluate_inventory_rules(df)

        self.assertEqual(
            [signal.id for signal in first],
            [signal.id for signal in second],
        )

        ranking = [
            (
                -signal.priority_rank,
                signal.id,
            )
            for signal in first
        ]

        self.assertEqual(
            ranking,
            sorted(ranking),
        )

    def test_todos_os_sinais_tem_evidencia(self):
        df = pd.DataFrame(
            [
                make_row(
                    Estoque_Atual=5,
                    Estoque_Minimo=10,
                    gap_estoque_minimo=-5,
                    razao_estoque_minimo=0.5,
                    cobertura_meses=0.5,
                )
            ]
        )

        signals = evaluate_inventory_rules(df)

        self.assertGreater(
            len(signals),
            0,
        )

        self.assertTrue(
            all(
                signal.evidences
                for signal in signals
            )
        )


class TestRealProjectData(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        estoque = pd.read_csv(
            DATA / "estoque_ficticio.csv",
            sep=SEP,
            encoding=ENC,
        )

        consumo = pd.read_csv(
            DATA / "consumo_ficticio.csv",
            sep=SEP,
            encoding=ENC,
        )

        cls.features, _ = build_material_features(
            estoque,
            consumo,
        )

        cls.signals = evaluate_inventory_rules(
            cls.features
        )

    def test_mat005_detectado_abaixo_do_minimo(self):
        matches = [
            signal
            for signal in self.signals
            if (
                signal.entity == "MAT005"
                and signal.type
                == "ESTOQUE_ABAIXO_MINIMO"
            )
        ]

        self.assertEqual(
            len(matches),
            1,
        )

    def test_cobertura_real_detecta_mat005_e_mat007(self):
        entities = {
            signal.entity
            for signal in self.signals
            if signal.type == "BAIXA_COBERTURA"
        }

        self.assertIn(
            "MAT005",
            entities,
        )

        self.assertIn(
            "MAT007",
            entities,
        )


if __name__ == "__main__":
    unittest.main()