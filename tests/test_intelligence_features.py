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


from intelligence.features import (  # noqa: E402
    build_consumption_features,
    build_material_features,
    build_stock_features,
)
from lobo_common import DATA, ENC, SEP  # noqa: E402


class BaseFeaturesTest(unittest.TestCase):
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


class TestStockFeatures(BaseFeaturesTest):
    def test_quantidade_real_de_materiais(self):
        result = build_stock_features(
            self.estoque
        )

        self.assertEqual(len(result), 20)

    def test_mat001_metricas_estoque(self):
        result = build_stock_features(
            self.estoque
        )

        row = result.loc[
            result["Material_ID"] == "MAT001"
        ].iloc[0]

        self.assertEqual(
            row["Estoque_Atual"],
            120,
        )

        self.assertEqual(
            row["Estoque_Minimo"],
            40,
        )

        self.assertEqual(
            row["gap_estoque_minimo"],
            80,
        )

        self.assertAlmostEqual(
            row["razao_estoque_minimo"],
            3.0,
        )

        self.assertAlmostEqual(
            row["percentual_gap_minimo"],
            200.0,
        )

        self.assertAlmostEqual(
            row["valor_estoque_atual"],
            1176.0,
        )

        self.assertAlmostEqual(
            row["valor_estoque_minimo"],
            392.0,
        )

    def test_mat005_abaixo_do_minimo(self):
        result = build_stock_features(
            self.estoque
        )

        row = result.loc[
            result["Material_ID"] == "MAT005"
        ].iloc[0]

        self.assertEqual(
            row["Estoque_Atual"],
            7,
        )

        self.assertEqual(
            row["Estoque_Minimo"],
            12,
        )

        self.assertEqual(
            row["gap_estoque_minimo"],
            -5,
        )

        self.assertAlmostEqual(
            row["razao_estoque_minimo"],
            7 / 12,
        )

        self.assertAlmostEqual(
            row["percentual_gap_minimo"],
            -41.66666666666667,
        )

        self.assertAlmostEqual(
            row["valor_estoque_atual"],
            672.0,
        )

        self.assertAlmostEqual(
            row["valor_estoque_minimo"],
            1152.0,
        )

    def test_divisao_segura_com_minimo_zero(self):
        estoque = self.estoque.copy()

        estoque.loc[
            estoque["Material_ID"] == "MAT003",
            "Estoque_Minimo",
        ] = 0

        result = build_stock_features(
            estoque
        )

        row = result.loc[
            result["Material_ID"] == "MAT003"
        ].iloc[0]

        self.assertTrue(
            pd.isna(
                row["razao_estoque_minimo"]
            )
        )

        self.assertTrue(
            pd.isna(
                row["percentual_gap_minimo"]
            )
        )

    def test_input_estoque_nao_e_modificado(self):
        original_columns = list(
            self.estoque.columns
        )

        build_stock_features(
            self.estoque
        )

        self.assertEqual(
            list(self.estoque.columns),
            original_columns,
        )

        self.assertNotIn(
            "gap_estoque_minimo",
            self.estoque.columns,
        )


class TestConsumptionFeatures(BaseFeaturesTest):
    def test_contexto_real_do_periodo(self):
        _, context = build_consumption_features(
            self.consumo
        )

        self.assertEqual(
            context.data_inicio,
            "2026-06-01",
        )

        self.assertEqual(
            context.data_fim,
            "2026-08-31",
        )

        self.assertEqual(
            context.meses_analisados,
            3,
        )

        self.assertEqual(
            context.total_movimentacoes,
            180,
        )

        self.assertAlmostEqual(
            context.total_quantidade,
            1647.0,
        )

        self.assertAlmostEqual(
            context.total_custo,
            24606.20,
            places=2,
        )

    def test_mat001_consumo_real(self):
        result, _ = build_consumption_features(
            self.consumo
        )

        row = result.loc[
            result["Material_ID"] == "MAT001"
        ].iloc[0]

        self.assertEqual(
            row["quantidade_consumida"],
            169,
        )

        self.assertAlmostEqual(
            row["custo_consumo"],
            1656.20,
            places=2,
        )

        self.assertEqual(
            row["movimentacoes"],
            13,
        )

        self.assertAlmostEqual(
            row["media_mensal_consumo"],
            56.333333333333336,
        )

        self.assertAlmostEqual(
            row["consumo_ultimo_mes"],
            5.0,
        )

        self.assertAlmostEqual(
            row["consumo_mes_anterior"],
            67.0,
        )

        self.assertAlmostEqual(
            row["variacao_consumo_pct"],
            -92.53731343283582,
        )

    def test_mat007_consumo_real(self):
        result, _ = build_consumption_features(
            self.consumo
        )

        row = result.loc[
            result["Material_ID"] == "MAT007"
        ].iloc[0]

        self.assertEqual(
            row["quantidade_consumida"],
            244,
        )

        self.assertAlmostEqual(
            row["custo_consumo"],
            3050.0,
        )

        self.assertEqual(
            row["movimentacoes"],
            14,
        )

        self.assertAlmostEqual(
            row["consumo_ultimo_mes"],
            47.0,
        )

        self.assertAlmostEqual(
            row["consumo_mes_anterior"],
            87.0,
        )

        self.assertAlmostEqual(
            row["variacao_consumo_pct"],
            -45.97701149425287,
        )

    def test_participacoes_totalizam_cem_porcento(self):
        result, _ = build_consumption_features(
            self.consumo
        )

        self.assertAlmostEqual(
            result[
                "participacao_consumo_pct"
            ].sum(),
            100.0,
        )

        self.assertAlmostEqual(
            result[
                "participacao_custo_pct"
            ].sum(),
            100.0,
        )

    def test_variacao_sem_base_nao_e_inventada(self):
        result, _ = build_consumption_features(
            self.consumo
        )

        row = result.loc[
            result["Material_ID"] == "MAT018"
        ].iloc[0]

        self.assertEqual(
            row["consumo_mes_anterior"],
            0.0,
        )

        self.assertTrue(
            pd.isna(
                row["variacao_consumo_pct"]
            )
        )

    def test_input_consumo_nao_e_modificado(self):
        original_columns = list(
            self.consumo.columns
        )

        build_consumption_features(
            self.consumo
        )

        self.assertEqual(
            list(self.consumo.columns),
            original_columns,
        )

        self.assertNotIn(
            "custo_movimentacao",
            self.consumo.columns,
        )


class TestMaterialFeatures(BaseFeaturesTest):
    def test_tabela_principal_tem_20_materiais(self):
        result, _ = build_material_features(
            self.estoque,
            self.consumo,
        )

        self.assertEqual(
            len(result),
            20,
        )

        self.assertEqual(
            result["Material_ID"].nunique(),
            20,
        )

    def test_mat005_cobertura_real(self):
        result, _ = build_material_features(
            self.estoque,
            self.consumo,
        )

        row = result.loc[
            result["Material_ID"] == "MAT005"
        ].iloc[0]

        self.assertAlmostEqual(
            row["cobertura_meses"],
            0.7241379310344828,
        )

    def test_mat007_cobertura_real(self):
        result, _ = build_material_features(
            self.estoque,
            self.consumo,
        )

        row = result.loc[
            result["Material_ID"] == "MAT007"
        ].iloc[0]

        self.assertAlmostEqual(
            row["cobertura_meses"],
            0.9098360655737705,
        )

    def test_contexto_preservado_na_integracao(self):
        _, context = build_material_features(
            self.estoque,
            self.consumo,
        )

        self.assertEqual(
            context.total_movimentacoes,
            180,
        )

        self.assertAlmostEqual(
            context.total_quantidade,
            1647.0,
        )

        self.assertAlmostEqual(
            context.total_custo,
            24606.20,
            places=2,
        )

    def test_coluna_obrigatoria_estoque_faltante_falha(self):
        estoque = self.estoque.drop(
            columns=["Lead_Time_Dias"]
        )

        with self.assertRaises(ValueError):
            build_material_features(
                estoque,
                self.consumo,
            )

    def test_coluna_obrigatoria_consumo_faltante_falha(self):
        consumo = self.consumo.drop(
            columns=["Quantidade"]
        )

        with self.assertRaises(ValueError):
            build_material_features(
                self.estoque,
                consumo,
            )


if __name__ == "__main__":
    unittest.main()