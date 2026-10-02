from __future__ import annotations

import math

import pandas as pd

from .models import (
    Evidence,
    Severity,
    Signal,
)


REQUIRED_COLUMNS = {
    "Material_ID",
    "Material",
    "custo_consumo",
    "participacao_custo_pct",
    "participacao_consumo_pct",
}


COST_CONCENTRATION_THRESHOLD_PCT = 10.0
COST_CONCENTRATION_HIGH_THRESHOLD_PCT = 15.0

COMBINED_COST_THRESHOLD_PCT = 10.0
COMBINED_CONSUMPTION_THRESHOLD_PCT = 10.0


def _validate_columns(
    features: pd.DataFrame,
) -> None:
    missing = REQUIRED_COLUMNS.difference(
        features.columns
    )

    if missing:
        missing_text = ", ".join(
            sorted(missing)
        )

        raise ValueError(
            "Colunas obrigatórias ausentes "
            "para regras de custo: "
            f"{missing_text}"
        )


def _safe_float(
    value: object,
) -> float | None:
    if value is None:
        return None

    try:
        number = float(value)
    except (TypeError, ValueError):
        return None

    if not math.isfinite(number):
        return None

    return number


def rule_cost_concentration(
    row: pd.Series,
    threshold_pct: float = (
        COST_CONCENTRATION_THRESHOLD_PCT
    ),
    high_threshold_pct: float = (
        COST_CONCENTRATION_HIGH_THRESHOLD_PCT
    ),
) -> Signal | None:
    """
    Identifica materiais com participação individual
    relevante no custo total de consumo.

    Limites padrão da V0.3.0:
    - >= 10%: ATENCAO;
    - >= 15%: ALTO.

    Os limites são explícitos e configuráveis.
    """

    material_id = str(
        row["Material_ID"]
    ).strip()

    material = str(
        row["Material"]
    ).strip()

    cost = _safe_float(
        row["custo_consumo"]
    )

    cost_share = _safe_float(
        row["participacao_custo_pct"]
    )

    if (
        cost is None
        or cost_share is None
    ):
        return None

    if cost <= 0:
        return None

    if cost_share < threshold_pct:
        return None

    severity = (
        Severity.ALTO
        if cost_share >= high_threshold_pct
        else Severity.ATENCAO
    )

    evidences = (
        Evidence(
            source="consumo",
            metric="custo_consumo",
            value=cost,
            description=(
                "Custo total estimado de consumo "
                "do material no período analisado."
            ),
        ),
        Evidence(
            source="consumo",
            metric="participacao_custo_pct",
            value=cost_share,
            description=(
                "Participação percentual do material "
                "no custo total de consumo."
            ),
        ),
    )

    return Signal(
        id=f"COST-CONCENTRATION-{material_id}",
        type="CONCENTRACAO_CUSTO",
        severity=severity,
        title=(
            f"Concentração relevante de custo — "
            f"{material}"
        ),
        description=(
            f"{material} representa "
            f"{cost_share:.2f}% do custo total "
            f"de consumo no período analisado."
        ),
        entity=material_id,
        rule_origin="rule_cost_concentration",
        evidences=evidences,
        metrics={
            "custo_consumo": cost,
            "participacao_custo_pct": cost_share,
            "limite_concentracao_pct": (
                threshold_pct
            ),
            "limite_alto_pct": (
                high_threshold_pct
            ),
        },
        recommended_action=(
            "Avaliar os principais direcionadores "
            "de custo do material e oportunidades "
            "de otimização de consumo, aquisição "
            "ou reposição."
        ),
    )


def rule_high_cost_and_consumption_impact(
    row: pd.Series,
    cost_threshold_pct: float = (
        COMBINED_COST_THRESHOLD_PCT
    ),
    consumption_threshold_pct: float = (
        COMBINED_CONSUMPTION_THRESHOLD_PCT
    ),
) -> Signal | None:
    """
    Identifica materiais simultaneamente relevantes
    em custo e volume consumido.

    O objetivo é diferenciar concentração puramente
    financeira de materiais que também possuem forte
    impacto operacional pelo volume.
    """

    material_id = str(
        row["Material_ID"]
    ).strip()

    material = str(
        row["Material"]
    ).strip()

    cost = _safe_float(
        row["custo_consumo"]
    )

    cost_share = _safe_float(
        row["participacao_custo_pct"]
    )

    consumption_share = _safe_float(
        row["participacao_consumo_pct"]
    )

    if (
        cost is None
        or cost_share is None
        or consumption_share is None
    ):
        return None

    if cost <= 0:
        return None

    if cost_share < cost_threshold_pct:
        return None

    if consumption_share < consumption_threshold_pct:
        return None

    evidences = (
        Evidence(
            source="consumo",
            metric="custo_consumo",
            value=cost,
            description=(
                "Custo total estimado de consumo "
                "do material no período analisado."
            ),
        ),
        Evidence(
            source="consumo",
            metric="participacao_custo_pct",
            value=cost_share,
            description=(
                "Participação percentual do material "
                "no custo total de consumo."
            ),
        ),
        Evidence(
            source="consumo",
            metric="participacao_consumo_pct",
            value=consumption_share,
            description=(
                "Participação percentual do material "
                "na quantidade total consumida."
            ),
        ),
    )

    return Signal(
        id=f"COST-CONSUMPTION-IMPACT-{material_id}",
        type="ALTO_IMPACTO_CUSTO_CONSUMO",
        severity=Severity.ALTO,
        title=(
            f"Alto impacto de custo e consumo — "
            f"{material}"
        ),
        description=(
            f"{material} concentra "
            f"{cost_share:.2f}% do custo total "
            f"e {consumption_share:.2f}% "
            f"da quantidade consumida."
        ),
        entity=material_id,
        rule_origin=(
            "rule_high_cost_and_consumption_impact"
        ),
        evidences=evidences,
        metrics={
            "custo_consumo": cost,
            "participacao_custo_pct": cost_share,
            "participacao_consumo_pct": (
                consumption_share
            ),
            "limite_custo_pct": (
                cost_threshold_pct
            ),
            "limite_consumo_pct": (
                consumption_threshold_pct
            ),
        },
        recommended_action=(
            "Priorizar análise deste material, "
            "pois ele combina impacto financeiro "
            "relevante com participação elevada "
            "no consumo operacional."
        ),
    )


def evaluate_cost_rules(
    features: pd.DataFrame,
) -> tuple[Signal, ...]:
    """
    Executa as regras determinísticas de inteligência
    de custos por material.

    A ordenação segue o padrão do projeto:
    severidade decrescente e ID crescente.
    """

    _validate_columns(
        features
    )

    signals: list[Signal] = []

    for _, row in features.iterrows():
        rules = (
            rule_cost_concentration,
            rule_high_cost_and_consumption_impact,
        )

        for rule in rules:
            signal = rule(
                row
            )

            if signal is not None:
                signals.append(
                    signal
                )

    return tuple(
        sorted(
            signals,
            key=lambda signal: (
                -signal.priority_rank,
                signal.id,
            ),
        )
    )