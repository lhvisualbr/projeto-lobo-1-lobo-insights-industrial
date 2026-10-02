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

    if math.isnan(number):
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
        # Abaixo do mínimo:
        # razão 1.0 = metade da pressão;
        # razão 0.0 = pressão máxima.
        normalized = (
            0.5
            + 0.5
            * _clamp(
                1.0 - ratio
            )
        )

    elif ratio <= NEAR_MINIMUM_RATIO:
        # Entre o mínimo e 125% do mínimo:
        # a pressão cai gradualmente até zero.
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


def score_signal(
    signal: Signal,
) -> PriorityScore:
    """
    Calcula score adicional de 0 a 100.

    Componentes máximos:
    - pressão de estoque: 30;
    - baixa cobertura: 25;
    - lead time: 20;
    - criticidade: 25.

    A severidade NÃO faz parte do score porque é utilizada
    como critério primário na ordenação.
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