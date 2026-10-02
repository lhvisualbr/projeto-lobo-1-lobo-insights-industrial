import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"

if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


from lobo_locale import (  # noqa: E402
    numeros_localizados_equivalentes,
    textos_localizados_equivalentes,
)


class TestNumerosLocalizados(unittest.TestCase):
    def test_decimal_ponto_e_virgula(self):
        self.assertTrue(
            numeros_localizados_equivalentes("9.2", "9,2")
        )

    def test_decimal_com_duas_casas(self):
        self.assertTrue(
            numeros_localizados_equivalentes("136.70", "136,70")
        )

    def test_milhar_e_decimal_en_us_pt_br(self):
        self.assertTrue(
            numeros_localizados_equivalentes(
                "24,606.20",
                "24.606,20",
            )
        )

    def test_milhar_ambiguo(self):
        self.assertTrue(
            numeros_localizados_equivalentes(
                "1,647",
                "1.647",
            )
        )

    def test_percentual(self):
        self.assertTrue(
            numeros_localizados_equivalentes(
                "25.7%",
                "25,7%",
            )
        )

    def test_valores_realmente_diferentes(self):
        self.assertFalse(
            numeros_localizados_equivalentes(
                "24,606.20",
                "24.606,21",
            )
        )


class TestTextosLocalizados(unittest.TestCase):
    def test_texto_com_decimal_localizado(self):
        self.assertTrue(
            textos_localizados_equivalentes(
                "Media de 9.2 unidades por movimentacao",
                "Media de 9,2 unidades por movimentacao",
            )
        )

    def test_texto_com_custo_localizado(self):
        self.assertTrue(
            textos_localizados_equivalentes(
                "do custo estimado total de 24,606.20",
                "do custo estimado total de 24.606,20",
            )
        )

    def test_texto_com_varios_numeros(self):
        self.assertTrue(
            textos_localizados_equivalentes(
                "Maior consumo: MECANICA - 423 un. (25.7%)",
                "Maior consumo: MECANICA - 423 un. (25,7%)",
            )
        )

    def test_texto_nao_numerico_diferente(self):
        self.assertFalse(
            textos_localizados_equivalentes(
                "Maior consumo: MECANICA - 423 un.",
                "Menor consumo: MECANICA - 423 un.",
            )
        )

    def test_valor_numerico_realmente_diferente(self):
        self.assertFalse(
            textos_localizados_equivalentes(
                "Custo 24,606.20",
                "Custo 24.606,21",
            )
        )

    def test_objetos_nao_textuais_diferentes(self):
        self.assertFalse(
            textos_localizados_equivalentes(
                100,
                101,
            )
        )

    def test_objetos_iguais(self):
        self.assertTrue(
            textos_localizados_equivalentes(
                100,
                100,
            )
        )


if __name__ == "__main__":
    unittest.main()