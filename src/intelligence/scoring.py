from __future__ import annotations

import math
import unicodedata
from dataclasses import dataclass
from typing import Any, Iterable

from intelligence.models import (
    Evidence,
    Signal,
)


STOCK_PRESSURE_WEIGHT = 30.0
COVERAGE_PRESSURE_WEIGHT = 25.0
LEAD_TIME_WEIGHT = 20.0
CRITICALITY_WEIGHT = 25.0

CONSUMPTION_VARIATION_WEIGHT = 25.0
CONSUMPTION_VARIATION_REFERENCE_PCT = 100.0

COST_SHARE_WEIGHT = 25.0
COST_SHARE_REFERENCE_PCT = 20.0

COST_CONSUMPTION_SHARE_WEIGHT = 20.0
COST_CONSUMPTION_SHARE_REFERENCE_PCT = 20.0

NEAR_MINIMUM_RATIO = 1.25
LEAD_TIME_REFERENCE_DAYS = 30.0


@dataclass(frozen=True, slots=True)
class PriorityScore:
    """
    Pontuação adicional para priorização.

    Não substitui a severidade do Signal.
    """

    total: float
    components: dict[str, float]
    explanations: tuple[str, ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "total": self.total,
            "components": dict(self.components),
            "explanations": list(self.explanations),
        }


@dataclass(frozen=True, slots=True)
class RankedSignal:
    """
    Signal acompanhado de sua pontuação de prioridade.
    """

    signal: Signal
    priority: PriorityScore

    def to_dict(self) -> dict[str, Any]:
        return {
            "signal": self.signal.to_dict(),
            "priority": self.priority.to_dict(),
        }


def _normalize_text(value: object) -> str:
    text = str(value).strip()

    normalized = unicodedata.normalize(
        "NFKD",
        text,
    )

    return "".join(
        char
        for char in normalized
        if not unicodedata.combining(char)
    ).casefold()


def _as_float(value: object) -> float | None:
    if value is None:
        return None

    try:
        number = float(value)
    except (TypeError, ValueError):
        return None

    if not math.isfinite(number):
        return None

    return number


def _clamp(
    value: float,
    minimum: float = 0.0,
    maximum: float = 1.0,
) -> float:
    return max(
        minimum,
        min(maximum, value),
    )


def _metric(
    signal: Signal,
    name: str,
) -> float | None:
    return _as_float(
        signal.metrics.get(name)
    )


def _find_evidence(
    signal: Signal,
    metric: str,
) -> Evidence | None:
    expected = _normalize_text(metric)

    for evidence in signal.evidences:
        if _normalize_text(
            evidence.metric
        ) == expected:
            return evidence

    return None


def _criticality_component(
    signal: Signal,
) -> tuple[float, str | None]:
    evidence = _find_evidence(
        signal,
        "Criticidade",
    )

    if evidence is None:
        return 0.0, None

    criticality = _normalize_text(
        evidence.value
    )

    if criticality == "alta":
        score = CRITICALITY_WEIGHT

    elif criticality == "media":
        score = CRITICALITY_WEIGHT * 0.5

    else:
        score = 0.0

    if score == 0:
        return 0.0, None

    return (
        score,
        (
            f"Criticidade {evidence.value}: "
            f"+{score:.2f} pontos."
        ),
    )


def _stock_pressure_component(
    signal: Signal,
) -> tuple[float, str | None]:
    atual = _metric(
        signal,
        "estoque_atual",
    )

    minimo = _metric(
        signal,
        "estoque_minimo",
    )

    ratio = _metric(
        signal,
        "razao_estoque_minimo",
    )

    if (
        ratio is None
        and atual is not None
        and minimo is not None
        and minimo > 0
    ):
        ratio = atual / minimo

    if ratio is None:
        return 0.0, None

    if ratio < 1.0:
        normalized = (
            0.5
            + 0.5
            * _clamp(
                1.0 - ratio
            )
        )

    elif ratio <= NEAR_MINIMUM_RATIO:
        normalized = (
            0.5
            * _clamp(
                (
                    NEAR_MINIMUM_RATIO
                    - ratio
                )
                / (
                    NEAR_MINIMUM_RATIO
                    - 1.0
                )
            )
        )

    else:
        normalized = 0.0

    score = (
        normalized
        * STOCK_PRESSURE_WEIGHT
    )

    if score == 0:
        return 0.0, None

    return (
        score,
        (
            f"Pressão de estoque "
            f"(razão atual/mínimo={ratio:.3f}): "
            f"+{score:.2f} pontos."
        ),
    )


def _coverage_component(
    signal: Signal,
) -> tuple[float, str | None]:
    coverage = _metric(
        signal,
        "cobertura_meses",
    )

    if coverage is None:
        return 0.0, None

    if coverage >= 1.0:
        return 0.0, None

    normalized = _clamp(
        1.0 - coverage
    )

    score = (
        normalized
        * COVERAGE_PRESSURE_WEIGHT
    )

    if score == 0:
        return 0.0, None

    return (
        score,
        (
            f"Cobertura estimada de "
            f"{coverage:.3f} mês(es): "
            f"+{score:.2f} pontos."
        ),
    )


def _lead_time_component(
    signal: Signal,
) -> tuple[float, str | None]:
    lead_time = _metric(
        signal,
        "lead_time_dias",
    )

    if lead_time is None:
        evidence = _find_evidence(
            signal,
            "Lead_Time_Dias",
        )

        if evidence is not None:
            lead_time = _as_float(
                evidence.value
            )

    if (
        lead_time is None
        or lead_time <= 0
    ):
        return 0.0, None

    normalized = _clamp(
        lead_time
        / LEAD_TIME_REFERENCE_DAYS
    )

    score = (
        normalized
        * LEAD_TIME_WEIGHT
    )

    if score == 0:
        return 0.0, None

    return (
        score,
        (
            f"Lead time de {lead_time:g} dia(s): "
            f"+{score:.2f} pontos."
        ),
    )


def _consumption_variation_component(
    signal: Signal,
) -> tuple[float, str | None]:
    signal_type = _normalize_text(
        signal.type
    )

    supported_types = {
        "crescimento_consumo",
        "reducao_consumo",
    }

    if signal_type not in supported_types:
        return 0.0, None

    variation = _metric(
        signal,
        "variacao_consumo_pct",
    )

    if variation is None:
        return 0.0, None

    magnitude = abs(
        variation
    )

    normalized = _clamp(
        magnitude
        / CONSUMPTION_VARIATION_REFERENCE_PCT
    )

    score = (
        normalized
        * CONSUMPTION_VARIATION_WEIGHT
    )

    if score == 0:
        return 0.0, None

    return (
        score,
        (
            f"Variação de consumo de "
            f"{variation:.2f}% "
            f"(magnitude {magnitude:.2f}%): "
            f"+{score:.2f} pontos."
        ),
    )


def _cost_share_component(
    signal: Signal,
) -> tuple[float, str | None]:
    """
    Pontua a concentração financeira individual
    de sinais do domínio de custos.
    """

    signal_type = _normalize_text(
        signal.type
    )

    supported_types = {
        "concentracao_custo",
        "alto_impacto_custo_consumo",
    }

    if signal_type not in supported_types:
        return 0.0, None

    cost_share = _metric(
        signal,
        "participacao_custo_pct",
    )

    if cost_share is None:
        return 0.0, None

    if cost_share <= 0:
        return 0.0, None

    normalized = _clamp(
        cost_share
        / COST_SHARE_REFERENCE_PCT
    )

    score = (
        normalized
        * COST_SHARE_WEIGHT
    )

    if score == 0:
        return 0.0, None

    return (
        score,
        (
            f"Participação no custo total de "
            f"{cost_share:.2f}%: "
            f"+{score:.2f} pontos."
        ),
    )


def _cost_consumption_share_component(
    signal: Signal,
) -> tuple[float, str | None]:
    """
    Acrescenta pressão de volume somente ao sinal
    combinado de alto impacto de custo e consumo.
    """

    signal_type = _normalize_text(
        signal.type
    )

    if signal_type != "alto_impacto_custo_consumo":
        return 0.0, None

    consumption_share = _metric(
        signal,
        "participacao_consumo_pct",
    )

    if consumption_share is None:
        return 0.0, None

    if consumption_share <= 0:
        return 0.0, None

    normalized = _clamp(
        consumption_share
        / COST_CONSUMPTION_SHARE_REFERENCE_PCT
    )

    score = (
        normalized
        * COST_CONSUMPTION_SHARE_WEIGHT
    )

    if score == 0:
        return 0.0, None

    return (
        score,
        (
            f"Participação no volume consumido de "
            f"{consumption_share:.2f}%: "
            f"+{score:.2f} pontos."
        ),
    )


def score_signal(
    signal: Signal,
) -> PriorityScore:
    """
    Calcula score adicional de prioridade.

    Estoque:
    - pressão de estoque: até 30;
    - baixa cobertura: até 25;
    - lead time: até 20;
    - criticidade: até 25.

    Comportamento de consumo:
    - magnitude da variação: até 25.

    Custos:
    - participação no custo total: até 25;
    - participação no volume, para sinal combinado:
      até 20.

    A severidade NÃO faz parte do score porque permanece
    como critério primário da ordenação.

    O total final permanece limitado entre 0 e 100.
    """

    components: dict[str, float] = {}
    explanations: list[str] = []

    calculators = (
        (
            "stock_pressure",
            _stock_pressure_component,
        ),
        (
            "coverage_pressure",
            _coverage_component,
        ),
        (
            "lead_time",
            _lead_time_component,
        ),
        (
            "criticality",
            _criticality_component,
        ),
    )

    for name, calculator in calculators:
        value, explanation = calculator(
            signal
        )

        components[name] = round(
            value,
            6,
        )

        if explanation is not None:
            explanations.append(
                explanation
            )

    signal_type = _normalize_text(
        signal.type
    )

    if signal_type in {
        "crescimento_consumo",
        "reducao_consumo",
    }:
        value, explanation = (
            _consumption_variation_component(
                signal
            )
        )

        components[
            "consumption_variation"
        ] = round(
            value,
            6,
        )

        if explanation is not None:
            explanations.append(
                explanation
            )

    if signal_type in {
        "concentracao_custo",
        "alto_impacto_custo_consumo",
    }:
        value, explanation = (
            _cost_share_component(
                signal
            )
        )

        components[
            "cost_share"
        ] = round(
            value,
            6,
        )

        if explanation is not None:
            explanations.append(
                explanation
            )

    if signal_type == "alto_impacto_custo_consumo":
        value, explanation = (
            _cost_consumption_share_component(
                signal
            )
        )

        components[
            "consumption_share"
        ] = round(
            value,
            6,
        )

        if explanation is not None:
            explanations.append(
                explanation
            )

    total = round(
        sum(components.values()),
        6,
    )

    total = min(
        100.0,
        max(0.0, total),
    )

    return PriorityScore(
        total=total,
        components=components,
        explanations=tuple(
            explanations
        ),
    )


def prioritize_signals(
    signals: Iterable[Signal],
) -> tuple[RankedSignal, ...]:
    """
    Ordenação determinística.

    Ordem:
    1. severidade decrescente;
    2. score decrescente;
    3. ID crescente.

    Portanto, o score nunca permite que um sinal ALTO
    ultrapasse um sinal CRITICO apenas pela pontuação.
    """

    ranked = [
        RankedSignal(
            signal=signal,
            priority=score_signal(
                signal
            ),
        )
        for signal in signals
    ]

    return tuple(
        sorted(
            ranked,
            key=lambda item: (
                -item.signal.priority_rank,
                -item.priority.total,
                item.signal.id,
            ),
        )
    )