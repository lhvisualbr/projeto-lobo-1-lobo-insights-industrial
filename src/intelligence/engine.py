from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Any

import pandas as pd

from intelligence.consumption_rules import (
    evaluate_consumption_rules,
)
from intelligence.cost_rules import (
    evaluate_cost_rules,
)
from intelligence.features import (
    FeatureContext,
    build_material_features,
)
from intelligence.lead_time_rules import (
    evaluate_lead_time_rules,
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
from intelligence.supplier_rules import (
    evaluate_supplier_rules,
)


SUPPLIER_SIGNAL_TYPES = {
    "CONCENTRACAO_FORNECEDOR",
    "EXPOSICAO_FORNECEDOR_CRITICO",
}


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


def _is_supplier_signal(
    signal: Signal,
) -> bool:
    return (
        signal.type
        in SUPPLIER_SIGNAL_TYPES
    )


def _build_indicators(
    features: pd.DataFrame,
    context: FeatureContext,
    signals: tuple[Signal, ...],
    ranked: tuple[RankedSignal, ...],
) -> dict[str, object]:
    counts = _count_signals(
        signals
    )

    material_entities = {
        signal.entity
        for signal in signals
        if not _is_supplier_signal(
            signal
        )
    }

    supplier_entities = {
        signal.entity
        for signal in signals
        if _is_supplier_signal(
            signal
        )
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
            len(material_entities)
        ),
        "fornecedores_com_sinal": int(
            len(supplier_entities)
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
    """
    Gera o resumo executivo textual sem alterar
    os cálculos ou a lógica de priorização.
    """

    def format_date_br(
        value: str | None,
    ) -> str | None:
        if not value:
            return None

        try:
            year, month, day = value.split("-")
            return f"{day}/{month}/{year}"
        except ValueError:
            return value

    def plural(
        quantity: int,
        singular: str,
        plural_form: str,
    ) -> str:
        return (
            singular
            if quantity == 1
            else plural_form
        )

    counts = _count_signals(
        signals
    )

    inicio = format_date_br(
        context.data_inicio
    )

    fim = format_date_br(
        context.data_fim
    )

    if inicio and fim:
        period = (
            f" entre {inicio} e {fim}"
        )
    else:
        period = ""

    total_signals = len(
        signals
    )

    signal_word = plural(
        total_signals,
        "sinal",
        "sinais",
    )

    criticos = counts["CRITICO"]
    altos = counts["ALTO"]
    atencao = counts["ATENCAO"]
    infos = counts["INFO"]

    critico_text = plural(
        criticos,
        "crítico",
        "críticos",
    )

    alto_text = plural(
        altos,
        "alto",
        "altos",
    )

    info_text = plural(
        infos,
        "informativo",
        "informativos",
    )

    base = (
        f"Foram analisadas "
        f"{context.total_movimentacoes} movimentações"
        f"{period}. "
        f"O motor identificou "
        f"{total_signals} {signal_word}: "
        f"{criticos} {critico_text}, "
        f"{altos} {alto_text}, "
        f"{atencao} de atenção e "
        f"{infos} {info_text}."
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
        + f"com score "
        + f"{top.priority.total:.2f} "
        + f"e severidade "
        + f"{top.signal.severity.value}."
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
        + evaluate_consumption_rules(
            features
        )
        + evaluate_cost_rules(
            features
        )
        + evaluate_lead_time_rules(
            features
        )
        + evaluate_supplier_rules(
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