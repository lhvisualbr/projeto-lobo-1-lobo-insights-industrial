from __future__ import annotations

import json
import math
from pathlib import Path
from typing import Any

from intelligence.engine import EngineResult


def _json_safe(value: Any) -> Any:
    """
    Converte valores para uma representação JSON segura.
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


def _format_number_ptbr(
    value: object,
    decimals: int = 2,
) -> str:
    """
    Formata números no padrão pt-BR.
    """

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


def _format_integer_ptbr(
    value: object,
) -> str:
    try:
        number = int(float(value))
    except (TypeError, ValueError):
        return str(value)

    return f"{number:,}".replace(",", ".")


def _escape_markdown_table(
    value: object,
) -> str:
    return (
        str(value)
        .replace("|", r"\|")
        .replace("\r", " ")
        .replace("\n", " ")
    )


def _group_recommendations(
    result: EngineResult,
) -> list[dict[str, object]]:
    """
    Agrupa recomendações por entidade para evitar repetição visual.

    A lógica original do engine permanece intacta.
    """

    groups: dict[str, dict[str, object]] = {}

    for recommendation in result.report.recommendations:
        entity = (
            recommendation.entity
            or "GERAL"
        )

        if entity not in groups:
            groups[entity] = {
                "entity": entity,
                "priority": recommendation.priority,
                "actions": [],
            }

        group = groups[entity]

        if (
            recommendation.priority.rank
            > group["priority"].rank
        ):
            group["priority"] = (
                recommendation.priority
            )

        actions = group["actions"]

        if recommendation.action not in actions:
            actions.append(
                recommendation.action
            )

    return sorted(
        groups.values(),
        key=lambda item: (
            -item["priority"].rank,
            str(item["entity"]),
        ),
    )


def result_to_dict(
    result: EngineResult,
) -> dict[str, Any]:
    return _json_safe(
        result.to_dict()
    )


def result_to_json(
    result: EngineResult,
    *,
    indent: int = 2,
) -> str:
    return json.dumps(
        result_to_dict(result),
        ensure_ascii=False,
        indent=indent,
        sort_keys=True,
        allow_nan=False,
    )


def result_to_markdown(
    result: EngineResult,
    *,
    max_signals: int = 10,
) -> str:
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
        (
            "- Materiais analisados: "
            + _format_integer_ptbr(
                indicators.get(
                    "materiais_analisados",
                    0,
                )
            )
        ),
        (
            "- Movimentações analisadas: "
            + _format_integer_ptbr(
                indicators.get(
                    "movimentacoes_analisadas",
                    0,
                )
            )
        ),
        (
            "- Quantidade total consumida: "
            + _format_integer_ptbr(
                indicators.get(
                    "quantidade_total_consumida",
                    0,
                )
            )
        ),
        (
            "- Custo total de consumo: R$ "
            + _format_number_ptbr(
                indicators.get(
                    "custo_total_consumo",
                    0,
                ),
                2,
            )
        ),
        (
            "- Total de sinais: "
            + str(
                report.total_signals
            )
        ),
        (
            "- Materiais com sinal: "
            + str(
                indicators.get(
                    "materiais_com_sinal",
                    0,
                )
            )
        ),
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
                        _format_number_ptbr(
                            item.priority.total,
                            2,
                        ),
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

    grouped = _group_recommendations(
        result
    )

    if not grouped:
        lines.append(
            "Nenhuma recomendação gerada pelas regras atuais."
        )
    else:
        for item in grouped:
            actions = " ".join(
                str(action)
                for action in item["actions"]
            )

            lines.append(
                f"- **{item['priority'].value} "
                f"[{item['entity']}]:** "
                f"{actions}"
            )

    lines.extend(
        [
            "",
            "## Rastreabilidade",
            "",
            (
                "- Período analisado: "
                f"{result.context.data_inicio or 'N/D'} "
                "a "
                f"{result.context.data_fim or 'N/D'}"
            ),
            (
                "- Meses analisados: "
                f"{result.context.meses_analisados}"
            ),
            (
                "- Movimentações analisadas: "
                f"{result.context.total_movimentacoes}"
            ),
            "",
            (
                "> As recomendações são suporte à decisão. "
                "Nenhuma ação operacional é executada automaticamente."
            ),
            "",
        ]
    )

    return "\n".join(lines)


def result_to_executive_text(
    result: EngineResult,
    *,
    max_signals: int = 5,
) -> str:
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
        lines.extend(
            [
                "",
                "PRINCIPAIS PRIORIDADES",
                "",
            ]
        )

        for position, item in enumerate(
            ranked,
            start=1,
        ):
            lines.append(
                f"{position}. "
                f"{item.signal.entity} — "
                f"{item.signal.title}"
            )

            lines.append(
                f"   Severidade: "
                f"{item.signal.severity.value}"
            )

            lines.append(
                f"   Score: "
                f"{_format_number_ptbr(item.priority.total, 2)}"
            )

            lines.append(
                f"   Motivo: "
                f"{item.signal.description}"
            )

    grouped = _group_recommendations(
        result
    )

    if grouped:
        lines.extend(
            [
                "",
                "RECOMENDAÇÕES",
                "",
            ]
        )

        for item in grouped:
            actions = " ".join(
                str(action)
                for action in item["actions"]
            )

            lines.append(
                f"- {item['entity']} "
                f"[{item['priority'].value}]: "
                f"{actions}"
            )

    return "\n".join(lines).strip()


def write_report_files(
    result: EngineResult,
    output_dir: str | Path,
    *,
    basename: str = "executive_intelligence_report",
) -> dict[str, Path]:
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