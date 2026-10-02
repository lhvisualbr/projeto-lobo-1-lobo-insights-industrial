from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


MetricValue = str | int | float | bool | None


class Severity(str, Enum):
    """
    Severidade de um sinal produzido pelo Executive Intelligence Engine.

    A ordem representa prioridade crescente.
    """

    INFO = "INFO"
    ATENCAO = "ATENCAO"
    ALTO = "ALTO"
    CRITICO = "CRITICO"

    @property
    def rank(self) -> int:
        return {
            Severity.INFO: 1,
            Severity.ATENCAO: 2,
            Severity.ALTO: 3,
            Severity.CRITICO: 4,
        }[self]


@dataclass(frozen=True, slots=True)
class Evidence:
    """
    Evidência objetiva que sustenta uma conclusão analítica.
    """

    source: str
    metric: str
    value: MetricValue
    description: str = ""

    def __post_init__(self) -> None:
        if not self.source.strip():
            raise ValueError("Evidence.source não pode ser vazio.")

        if not self.metric.strip():
            raise ValueError("Evidence.metric não pode ser vazio.")

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": self.source,
            "metric": self.metric,
            "value": self.value,
            "description": self.description,
        }


@dataclass(frozen=True, slots=True)
class Recommendation:
    """
    Recomendação de suporte à decisão.

    Não representa execução automática de uma ação.
    """

    action: str
    rationale: str
    entity: str | None = None
    priority: Severity = Severity.INFO

    def __post_init__(self) -> None:
        if not self.action.strip():
            raise ValueError(
                "Recommendation.action não pode ser vazio."
            )

        if not self.rationale.strip():
            raise ValueError(
                "Recommendation.rationale não pode ser vazio."
            )

    def to_dict(self) -> dict[str, Any]:
        return {
            "action": self.action,
            "rationale": self.rationale,
            "entity": self.entity,
            "priority": self.priority.value,
        }


@dataclass(frozen=True, slots=True)
class Signal:
    """
    Sinal analítico detectado pelo motor de inteligência.
    """

    id: str
    type: str
    severity: Severity
    title: str
    description: str
    entity: str
    rule_origin: str

    evidences: tuple[Evidence, ...] = ()
    metrics: dict[str, MetricValue] = field(
        default_factory=dict
    )
    recommended_action: str | None = None

    def __post_init__(self) -> None:
        required = {
            "id": self.id,
            "type": self.type,
            "title": self.title,
            "description": self.description,
            "entity": self.entity,
            "rule_origin": self.rule_origin,
        }

        for field_name, value in required.items():
            if not value.strip():
                raise ValueError(
                    f"Signal.{field_name} não pode ser vazio."
                )

        if not self.evidences:
            raise ValueError(
                "Todo Signal deve possuir pelo menos uma Evidence."
            )

    @property
    def priority_rank(self) -> int:
        return self.severity.rank

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "type": self.type,
            "severity": self.severity.value,
            "title": self.title,
            "description": self.description,
            "entity": self.entity,
            "evidences": [
                evidence.to_dict()
                for evidence in self.evidences
            ],
            "metrics": dict(self.metrics),
            "recommended_action": self.recommended_action,
            "rule_origin": self.rule_origin,
        }


@dataclass(frozen=True, slots=True)
class ExecutiveReport:
    """
    Resultado estruturado produzido pelo motor executivo.
    """

    signals: tuple[Signal, ...] = ()
    recommendations: tuple[Recommendation, ...] = ()
    indicators: dict[str, MetricValue] = field(
        default_factory=dict
    )
    summary: str = ""

    @property
    def total_signals(self) -> int:
        return len(self.signals)

    def count_by_severity(self) -> dict[str, int]:
        result = {
            severity.value: 0
            for severity in Severity
        }

        for signal in self.signals:
            result[signal.severity.value] += 1

        return result

    def ordered_signals(self) -> tuple[Signal, ...]:
        """
        Retorna os sinais por severidade decrescente.

        O ID funciona como critério secundário determinístico.
        """
        return tuple(
            sorted(
                self.signals,
                key=lambda signal: (
                    -signal.priority_rank,
                    signal.id,
                ),
            )
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "summary": self.summary,
            "total_signals": self.total_signals,
            "signals_by_severity": self.count_by_severity(),
            "indicators": dict(self.indicators),
            "signals": [
                signal.to_dict()
                for signal in self.ordered_signals()
            ],
            "recommendations": [
                recommendation.to_dict()
                for recommendation in self.recommendations
            ],
        }