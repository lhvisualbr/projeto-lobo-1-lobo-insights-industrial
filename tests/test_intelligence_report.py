import json
import sys
import tempfile
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


from intelligence.engine import run_engine  # noqa: E402
from intelligence.report import (  # noqa: E402
    result_to_dict,
    result_to_executive_text,
    result_to_json,
    result_to_markdown,
    write_report_files,
)
from lobo_common import DATA, ENC, SEP  # noqa: E402


class TestIntelligenceReport(unittest.TestCase):
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

        cls.result = run_engine(
            estoque,
            consumo,
        )

    def test_result_to_dict_possui_estrutura_esperada(self):
        data = result_to_dict(
            self.result
        )

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

    def test_json_e_valido(self):
        content = result_to_json(
            self.result
        )

        data = json.loads(
            content
        )

        self.assertEqual(
            data,
            result_to_dict(
                self.result
            ),
        )

    def test_json_e_deterministico(self):
        first = result_to_json(
            self.result
        )

        second = result_to_json(
            self.result
        )

        self.assertEqual(
            first,
            second,
        )

    def test_json_nao_emite_nan_ou_infinity(self):
        content = result_to_json(
            self.result
        )

        self.assertNotIn(
            "NaN",
            content,
        )

        self.assertNotIn(
            "Infinity",
            content,
        )

    def test_markdown_possui_secoes_executivas(self):
        content = result_to_markdown(
            self.result
        )

        self.assertIn(
            "# Lobo Insights Industrial — Executive Intelligence Report",
            content,
        )

        self.assertIn(
            "## Resumo executivo",
            content,
        )

        self.assertIn(
            "## Indicadores",
            content,
        )

        self.assertIn(
            "## Sinais por severidade",
            content,
        )

        self.assertIn(
            "## Principais prioridades",
            content,
        )

        self.assertIn(
            "## Recomendações",
            content,
        )

        self.assertIn(
            "## Rastreabilidade",
            content,
        )

    def test_markdown_respeita_limite_de_sinais(self):
        content = result_to_markdown(
            self.result,
            max_signals=1,
        )

        priorities = (
            content
            .split(
                "## Principais prioridades",
                1,
            )[1]
            .split(
                "## Recomendações",
                1,
            )[0]
        )

        data_rows = [
            line
            for line in priorities.splitlines()
            if (
                line.startswith("| ")
                and len(line) > 2
                and line[2].isdigit()
            )
        ]

        self.assertEqual(
            len(data_rows),
            1,
        )

    def test_markdown_limite_negativo_falha(self):
        with self.assertRaises(
            ValueError
        ):
            result_to_markdown(
                self.result,
                max_signals=-1,
            )

    def test_texto_executivo_possui_prioridades(self):
        content = result_to_executive_text(
            self.result,
            max_signals=3,
        )

        self.assertIn(
            self.result.report.summary,
            content,
        )

        self.assertIn(
            "Prioridades:",
            content,
        )

        top = self.result.ranked_signals[0]

        self.assertIn(
            top.signal.entity,
            content,
        )

        self.assertIn(
            top.signal.severity.value,
            content,
        )

    def test_texto_executivo_limite_negativo_falha(self):
        with self.assertRaises(
            ValueError
        ):
            result_to_executive_text(
                self.result,
                max_signals=-1,
            )

    def test_write_report_files_cria_tres_arquivos(self):
        with tempfile.TemporaryDirectory(
            prefix="lobo_report_test_"
        ) as tmp:
            files = write_report_files(
                self.result,
                tmp,
                basename="teste_executivo",
            )

            self.assertEqual(
                set(files),
                {
                    "json",
                    "markdown",
                    "text",
                },
            )

            self.assertTrue(
                all(
                    path.is_file()
                    for path in files.values()
                )
            )

    def test_arquivos_gravados_correspondem_ao_resultado(self):
        with tempfile.TemporaryDirectory(
            prefix="lobo_report_content_"
        ) as tmp:
            files = write_report_files(
                self.result,
                tmp,
                basename="relatorio",
            )

            json_content = files[
                "json"
            ].read_text(
                encoding="utf-8"
            )

            markdown_content = files[
                "markdown"
            ].read_text(
                encoding="utf-8"
            )

            text_content = files[
                "text"
            ].read_text(
                encoding="utf-8"
            )

            self.assertEqual(
                json.loads(json_content),
                result_to_dict(
                    self.result
                ),
            )

            self.assertEqual(
                markdown_content,
                result_to_markdown(
                    self.result
                ),
            )

            self.assertEqual(
                text_content.rstrip("\n"),
                result_to_executive_text(
                    self.result
                ),
            )

    def test_saida_utf8_preserva_acentos(self):
        markdown = result_to_markdown(
            self.result
        )

        executive = result_to_executive_text(
            self.result
        )

        self.assertIn(
            "Movimentações",
            markdown,
        )

        self.assertIn(
            "Recomendações",
            markdown,
        )

        self.assertIn(
            "movimentações",
            executive.lower(),
        )


if __name__ == "__main__":
    unittest.main()