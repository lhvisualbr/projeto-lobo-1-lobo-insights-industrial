from __future__ import annotations

import sys
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
from intelligence.report import write_report_files  # noqa: E402
from lobo_common import DATA, ENC, SEP  # noqa: E402


OUTPUT_DIR = ROOT / "reports" / "v0_3"
BASENAME = "executive_intelligence_report"


def format_number_ptbr(
    value: object,
    decimals: int = 2,
) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)

    formatted = f"{number:,.{decimals}f}"

    return (
        formatted
        .replace(",", "#")
        .replace(".", ",")
        .replace("#", ".")
    )


def format_integer_ptbr(
    value: object,
) -> str:
    try:
        number = int(float(value))
    except (TypeError, ValueError):
        return str(value)

    return f"{number:,}".replace(",", ".")


def carregar_dados() -> tuple[pd.DataFrame, pd.DataFrame]:
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

    return estoque, consumo


def main() -> int:
    print("=" * 72)
    print("LOBO INSIGHTS INDUSTRIAL — EXECUTIVE INTELLIGENCE V0.3.0")
    print("=" * 72)

    print()
    print("Carregando dados...")

    estoque, consumo = carregar_dados()

    print(
        f"Estoque: {len(estoque)} materiais"
    )

    print(
        f"Consumo: {len(consumo)} movimentações"
    )

    print()
    print("Executando Executive Intelligence Engine...")

    result = run_engine(
        estoque,
        consumo,
    )

    print("Engine concluído.")

    print()
    print("Gerando relatórios...")

    files = write_report_files(
        result,
        OUTPUT_DIR,
        basename=BASENAME,
    )

    print()
    print("=" * 72)
    print("RESUMO EXECUTIVO")
    print("=" * 72)
    print(result.report.summary)

    print()
    print("=" * 72)
    print("INDICADORES")
    print("=" * 72)

    indicators = result.report.indicators

    print(
        "materiais_analisados:",
        format_integer_ptbr(
            indicators["materiais_analisados"]
        ),
    )

    print(
        "movimentacoes_analisadas:",
        format_integer_ptbr(
            indicators["movimentacoes_analisadas"]
        ),
    )

    print(
        "quantidade_total_consumida:",
        format_integer_ptbr(
            indicators["quantidade_total_consumida"]
        ),
    )

    print(
        "custo_total_consumo: R$",
        format_number_ptbr(
            indicators["custo_total_consumo"],
            2,
        ),
    )

    print(
        "total_sinais:",
        indicators["total_sinais"],
    )

    print(
        "materiais_com_sinal:",
        indicators["materiais_com_sinal"],
    )

    print(
        "sinais_info:",
        indicators["sinais_info"],
    )

    print(
        "sinais_atencao:",
        indicators["sinais_atencao"],
    )

    print(
        "sinais_alto:",
        indicators["sinais_alto"],
    )

    print(
        "sinais_critico:",
        indicators["sinais_critico"],
    )

    print(
        "maior_priority_score:",
        format_number_ptbr(
            indicators["maior_priority_score"],
            2,
        ),
    )

    print()
    print("=" * 72)
    print("TOP 5 PRIORIDADES")
    print("=" * 72)

    if not result.ranked_signals:
        print(
            "Nenhum sinal identificado."
        )
    else:
        for position, item in enumerate(
            result.ranked_signals[:5],
            start=1,
        ):
            print(
                f"{position}. "
                f"{item.signal.entity} | "
                f"{item.signal.severity.value} | "
                f"score "
                f"{format_number_ptbr(item.priority.total, 2)}"
            )

            print(
                f"   {item.signal.title}"
            )

    print()
    print("=" * 72)
    print("ARQUIVOS GERADOS")
    print("=" * 72)

    for file_type, path in files.items():
        print(
            f"{file_type}: {path}"
        )

    print()
    print(
        "Total de sinais:",
        result.report.total_signals,
    )

    print(
        "Total de recomendações:",
        len(result.report.recommendations),
    )

    print()
    print("GERAÇÃO FINALIZADA COM SUCESSO")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())