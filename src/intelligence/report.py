from __future__ import annotations

import json
import math
from pathlib import Path
from typing import Any

from intelligence.engine import EngineResult


def _json_safe(value: Any) -> Any:
    """
    Converte valores para uma representação JSON segura.

    NaN e infinito não são emitidos como JSON inválido;
    nesses casos o valor passa a ser None.
    """

    if isinstance(value, dict):
        return {
            str(key): _json_safe(item)
            for key, item in value.items()
        }

    if isinstance(value, (list, tuple)):
        return [
            _json_safe(item)
            for item in value
        ]

    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None

    return value


def result_to_dict(
    result: EngineResult,
) -> dict[str, Any]:
    """
    Retorna a saída estruturada do engine pronta para serialização.
    """

    return _json_safe(
        result.to_dict()
    )


def result_to_json(
    result: EngineResult,
    *,
    indent: int = 2,
) -> str:
    """
    Gera JSON determinístico, legível e compatível com UTF-8.
    """

    return json.dumps(
        result_to_dict(result),
        ensure_ascii=False,
        indent=indent,
        sort_keys=True,
        allow_nan=False,
    )


def _escape_markdown_table(
    value: object,
) -> str:
    text = str(value)

    return (
        text
        .replace("|", r"\|")
        .replace("\r", " ")
        .replace("\n", " ")
    )


def result_to_markdown(
    result: EngineResult,
    *,
    max_signals: int = 10,
) -> str:
    """
    Gera relatório executivo em Markdown.

    A função apenas apresenta informações já produzidas
    pelo Executive Intelligence Engine.
    """

    if max_signals < 0:
        raise ValueError(
            "max_signals não pode ser negativo."
        )

    report = result.report
    indicators = report.indicators
    counts = report.count_by_severity()

    lines: list[str] = [
        "# Lobo Insights Industrial — Executive Intelligence Report",
        "",
        "## Resumo executivo",
        "",
        report.summary.strip(),
        "",
        "## Indicadores",
        "",
        f"- Materiais analisados: "
        f"{indicators.get('materiais_analisados', 0)}",
        f"- Movimentações analisadas: "
        f"{indicators.get('movimentacoes_analisadas', 0)}",
        f"- Quantidade total consumida: "
        f"{indicators.get('quantidade_total_consumida', 0)}",
        f"- Custo total de consumo: "
        f"{indicators.get('custo_total_consumo', 0)}",
        f"- Total de sinais: "
        f"{report.total_signals}",
        f"- Materiais com sinal: "
        f"{indicators.get('materiais_com_sinal', 0)}",
        "",
        "## Sinais por severidade",
        "",
        f"- CRITICO: {counts['CRITICO']}",
        f"- ALTO: {counts['ALTO']}",
        f"- ATENCAO: {counts['ATENCAO']}",
        f"- INFO: {counts['INFO']}",
        "",
        "## Principais prioridades",
        "",
    ]

    ranked = result.ranked_signals[
        :max_signals
    ]

    if not ranked:
        lines.append(
            "Nenhum sinal identificado pelas regras atuais."
        )
    else:
        lines.extend(
            [
                "| # | Entidade | Severidade | Score | Sinal |",
                "|---:|---|---|---:|---|",
            ]
        )

        for position, item in enumerate(
            ranked,
            start=1,
        ):
            lines.append(
                "| "
                + " | ".join(
                    [
                        str(position),
                        _escape_markdown_table(
                            item.signal.entity
                        ),
                        item.signal.severity.value,
                        f"{item.priority.total:.2f}",
                        _escape_markdown_table(
                            item.signal.title
                        ),
                    ]
                )
                + " |"
            )

    lines.extend(
        [
            "",
            "## Recomendações",
            "",
        ]
    )

    if not report.recommendations:
        lines.append(
            "Nenhuma recomendação gerada pelas regras atuais."
        )
    else:
        for recommendation in report.recommendations:
            entity = (
                f" [{recommendation.entity}]"
                if recommendation.entity
                else ""
            )

            lines.append(
                f"- **{recommendation.priority.value}"
                f"{entity}:** "
                f"{recommendation.action} "
                f"Motivo: {recommendation.rationale}"
            )

    lines.extend(
        [
            "",
            "## Rastreabilidade",
            "",
            f"- Período analisado: "
            f"{result.context.data_inicio or 'N/D'} "
            f"a "
            f"{result.context.data_fim or 'N/D'}",
            f"- Meses analisados: "
            f"{result.context.meses_analisados}",
            f"- Movimentações analisadas: "
            f"{result.context.total_movimentacoes}",
            "",
            "> As recomendações são suporte à decisão. "
            "Nenhuma ação operacional é executada automaticamente.",
            "",
        ]
    )

    return "\n".join(lines)


def result_to_executive_text(
    result: EngineResult,
    *,
    max_signals: int = 5,
) -> str:
    """
    Gera versão textual curta para leitura executiva.
    """

    if max_signals < 0:
        raise ValueError(
            "max_signals não pode ser negativo."
        )

    lines = [
        result.report.summary.strip(),
    ]

    ranked = result.ranked_signals[
        :max_signals
    ]

    if ranked:
        lines.append("")
        lines.append("Prioridades:")

        for position, item in enumerate(
            ranked,
            start=1,
        ):
            lines.append(
                f"{position}. "
                f"{item.signal.entity} — "
                f"{item.signal.title} | "
                f"{item.signal.severity.value} | "
                f"score {item.priority.total:.2f}"
            )

    if result.report.recommendations:
        lines.append("")
        lines.append("Recomendações:")

        for recommendation in result.report.recommendations:
            entity = (
                f"{recommendation.entity}: "
                if recommendation.entity
                else ""
            )

            lines.append(
                f"- {entity}"
                f"{recommendation.action}"
            )

    return "\n".join(lines).strip()


def write_report_files(
    result: EngineResult,
    output_dir: str | Path,
    *,
    basename: str = "executive_intelligence_report",
) -> dict[str, Path]:
    """
    Grava JSON, Markdown e TXT em UTF-8.

    Retorna os caminhos efetivamente criados.
    """

    if not basename.strip():
        raise ValueError(
            "basename não pode ser vazio."
        )

    destination = Path(
        output_dir
    )

    destination.mkdir(
        parents=True,
        exist_ok=True,
    )

    files = {
        "json": destination / f"{basename}.json",
        "markdown": destination / f"{basename}.md",
        "text": destination / f"{basename}.txt",
    }

    files["json"].write_text(
        result_to_json(result) + "\n",
        encoding="utf-8",
    )

    files["markdown"].write_text(
        result_to_markdown(result),
        encoding="utf-8",
    )

    files["text"].write_text(
        result_to_executive_text(result) + "\n",
        encoding="utf-8",
    )

    return files