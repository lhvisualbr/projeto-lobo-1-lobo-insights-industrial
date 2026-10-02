from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Any

import pandas as pd

from intelligence.features import (
    FeatureContext,
    build_material_features,
)
from intelligence.models import (
    ExecutiveReport,
    Recommendation,
    Severity,
    Signal,
)
from intelligence.rules import (
    evaluate_inventory_rules,
)
from intelligence.scoring import (
    RankedSignal,
    prioritize_signals,
)


@dataclass(frozen=True, slots=True)
class EngineResult:
    """
    Resultado completo do Executive Intelligence Engine.

    Mantém:
    - relatório executivo;
    - ranking detalhado;
    - features calculadas;
    - contexto do período.
    """

    report: ExecutiveReport
    ranked_signals: tuple[RankedSignal, ...]
    features: pd.DataFrame
    context: FeatureContext

    def to_dict(self) -> dict[str, Any]:
        return {
            "report": self.report.to_dict(),
            "ranking": [
                item.to_dict()
                for item in self.ranked_signals
            ],
            "context": self.context.to_dict(),
        }


def _signals_with_priority_score(
    ranked: tuple[RankedSignal, ...],
) -> tuple[Signal, ...]:
    """
    Cria cópias dos sinais com o score registrado
    nas métricas para rastreabilidade.

    O Signal original não é modificado.
    """

    result: list[Signal] = []

    for item in ranked:
        metrics = dict(
            item.signal.metrics
        )

        metrics["priority_score"] = (
            item.priority.total
        )

        result.append(
            replace(
                item.signal,
                metrics=metrics,
            )
        )

    return tuple(result)


def _build_recommendations(
    ranked: tuple[RankedSignal, ...],
) -> tuple[Recommendation, ...]:
    """
    Converte ações sugeridas pelos sinais em recomendações.

    Remove duplicações exatas de entidade + ação
    preservando a ordem de prioridade.
    """

    recommendations: list[Recommendation] = []
    seen: set[tuple[str, str]] = set()

    for item in ranked:
        signal = item.signal
        action = signal.recommended_action

        if not action:
            continue

        key = (
            signal.entity,
            action,
        )

        if key in seen:
            continue

        seen.add(key)

        recommendations.append(
            Recommendation(
                action=action,
                rationale=signal.description,
                entity=signal.entity,
                priority=signal.severity,
            )
        )

    return tuple(
        recommendations
    )


def _count_signals(
    signals: tuple[Signal, ...],
) -> dict[str, int]:
    result = {
        severity.value: 0
        for severity in Severity
    }

    for signal in signals:
        result[
            signal.severity.value
        ] += 1

    return result


def _build_indicators(
    features: pd.DataFrame,
    context: FeatureContext,
    signals: tuple[Signal, ...],
    ranked: tuple[RankedSignal, ...],
) -> dict[str, object]:
    counts = _count_signals(
        signals
    )

    entities = {
        signal.entity
        for signal in signals
    }

    top_score = (
        ranked[0].priority.total
        if ranked
        else 0.0
    )

    return {
        "materiais_analisados": int(
            len(features)
        ),
        "movimentacoes_analisadas": int(
            context.total_movimentacoes
        ),
        "quantidade_total_consumida": float(
            context.total_quantidade
        ),
        "custo_total_consumo": float(
            context.total_custo
        ),
        "total_sinais": int(
            len(signals)
        ),
        "materiais_com_sinal": int(
            len(entities)
        ),
        "sinais_info": int(
            counts["INFO"]
        ),
        "sinais_atencao": int(
            counts["ATENCAO"]
        ),
        "sinais_alto": int(
            counts["ALTO"]
        ),
        "sinais_critico": int(
            counts["CRITICO"]
        ),
        "maior_priority_score": float(
            top_score
        ),
    }


def _build_summary(
    context: FeatureContext,
    signals: tuple[Signal, ...],
    ranked: tuple[RankedSignal, ...],
) -> str:
    counts = _count_signals(
        signals
    )

    period = ""

    if (
        context.data_inicio
        and context.data_fim
    ):
        period = (
            f" no período de "
            f"{context.data_inicio} a "
            f"{context.data_fim}"
        )

    base = (
        f"Foram analisadas "
        f"{context.total_movimentacoes} movimentações"
        f"{period}. "
        f"O motor identificou {len(signals)} sinal(is): "
        f"{counts['CRITICO']} crítico(s), "
        f"{counts['ALTO']} alto(s), "
        f"{counts['ATENCAO']} de atenção e "
        f"{counts['INFO']} informativo(s)."
    )

    if not ranked:
        return (
            base
            + " Nenhuma condição de risco "
            + "foi identificada pelas regras atuais."
        )

    top = ranked[0]

    return (
        base
        + " A maior prioridade identificada foi "
        + f"{top.signal.entity} — "
        + f"{top.signal.title}, "
        + f"com score {top.priority.total:.2f} "
        + f"e severidade {top.signal.severity.value}."
    )


def run_engine(
    estoque: pd.DataFrame,
    consumo: pd.DataFrame,
) -> EngineResult:
    """
    Executa o fluxo completo da inteligência executiva.

    Fluxo:
        dados
        -> features
        -> regras
        -> sinais
        -> scoring/priorização
        -> recomendações
        -> ExecutiveReport

    Nenhum dado de entrada é alterado.
    """

    features, context = (
        build_material_features(
            estoque,
            consumo,
        )
    )

    raw_signals = (
        evaluate_inventory_rules(
            features
        )
    )

    ranked = prioritize_signals(
        raw_signals
    )

    report_signals = (
        _signals_with_priority_score(
            ranked
        )
    )

    recommendations = (
        _build_recommendations(
            ranked
        )
    )

    indicators = _build_indicators(
        features,
        context,
        report_signals,
        ranked,
    )

    summary = _build_summary(
        context,
        report_signals,
        ranked,
    )

    report = ExecutiveReport(
        signals=report_signals,
        recommendations=recommendations,
        indicators=indicators,
        summary=summary,
    )

    return EngineResult(
        report=report,
        ranked_signals=ranked,
        features=features,
        context=context,
    )