from __future__ import annotations

import math

import pandas as pd

from .models import Evidence, Severity, Signal


REQUIRED_COLUMNS = {
    "Material_ID",
    "Material",
    "consumo_mes_anterior",
    "consumo_ultimo_mes",
    "variacao_consumo_pct",
}


def _validate_columns(features: pd.DataFrame) -> None:
    missing = REQUIRED_COLUMNS.difference(features.columns)

    if missing:
        missing_text = ", ".join(sorted(missing))
        raise ValueError(
            "Colunas obrigatórias ausentes para regras de consumo: "
            f"{missing_text}"
        )


def _safe_float(value: object) -> float | None:
    if value is None:
        return None

    try:
        number = float(value)
    except (TypeError, ValueError):
        return None

    if not math.isfinite(number):
        return None

    return number


def rule_consumption_growth(
    row: pd.Series,
    threshold_pct: float = 30.0,
) -> Signal | None:
    """
    Identifica crescimento relevante de consumo entre
    o mês anterior e o último mês disponível.

    A regra só é aplicável quando existe consumo positivo
    no mês anterior, evitando interpretar crescimento a
    partir de uma base igual a zero.
    """

    material_id = str(row["Material_ID"]).strip()
    material = str(row["Material"]).strip()

    previous = _safe_float(
        row["consumo_mes_anterior"]
    )
    latest = _safe_float(
        row["consumo_ultimo_mes"]
    )
    variation = _safe_float(
        row["variacao_consumo_pct"]
    )

    if (
        previous is None
        or latest is None
        or variation is None
    ):
        return None

    if previous <= 0:
        return None

    if latest <= previous:
        return None

    if variation < threshold_pct:
        return None

    severity = (
        Severity.ALTO
        if variation >= 75.0
        else Severity.ATENCAO
    )

    evidences = (
        Evidence(
            source="consumo",
            metric="consumo_mes_anterior",
            value=previous,
            description=(
                "Quantidade consumida no mês anterior."
            ),
        ),
        Evidence(
            source="consumo",
            metric="consumo_ultimo_mes",
            value=latest,
            description=(
                "Quantidade consumida no último mês."
            ),
        ),
        Evidence(
            source="consumo",
            metric="variacao_consumo_pct",
            value=variation,
            description=(
                "Variação percentual entre o mês anterior "
                "e o último mês."
            ),
        ),
    )

    return Signal(
        id=f"CONSUMPTION-GROWTH-{material_id}",
        type="CRESCIMENTO_CONSUMO",
        severity=severity,
        title=f"Crescimento relevante de consumo — {material}",
        description=(
            f"{material} apresentou crescimento de "
            f"{variation:.2f}% no consumo do último mês, "
            f"passando de {previous:.2f} para {latest:.2f}."
        ),
        entity=material_id,
        rule_origin="rule_consumption_growth",
        evidences=evidences,
        metrics={
            "consumo_mes_anterior": previous,
            "consumo_ultimo_mes": latest,
            "variacao_consumo_pct": variation,
            "limite_variacao_pct": threshold_pct,
        },
        recommended_action=(
            "Verificar a causa do aumento de consumo e avaliar "
            "se o planejamento de reposição continua adequado."
        ),
    )


def rule_consumption_drop(
    row: pd.Series,
    threshold_pct: float = -30.0,
) -> Signal | None:
    """
    Identifica redução relevante de consumo.

    A redução não é tratada automaticamente como problema.
    O sinal indica uma mudança relevante no padrão histórico
    que pode justificar investigação operacional.
    """

    material_id = str(row["Material_ID"]).strip()
    material = str(row["Material"]).strip()

    previous = _safe_float(
        row["consumo_mes_anterior"]
    )
    latest = _safe_float(
        row["consumo_ultimo_mes"]
    )
    variation = _safe_float(
        row["variacao_consumo_pct"]
    )

    if (
        previous is None
        or latest is None
        or variation is None
    ):
        return None

    if previous <= 0:
        return None

    if latest >= previous:
        return None

    if variation > threshold_pct:
        return None

    severity = (
        Severity.ATENCAO
        if variation <= -50.0
        else Severity.INFO
    )

    evidences = (
        Evidence(
            source="consumo",
            metric="consumo_mes_anterior",
            value=previous,
            description=(
                "Quantidade consumida no mês anterior."
            ),
        ),
        Evidence(
            source="consumo",
            metric="consumo_ultimo_mes",
            value=latest,
            description=(
                "Quantidade consumida no último mês."
            ),
        ),
        Evidence(
            source="consumo",
            metric="variacao_consumo_pct",
            value=variation,
            description=(
                "Variação percentual entre o mês anterior "
                "e o último mês."
            ),
        ),
    )

    return Signal(
        id=f"CONSUMPTION-DROP-{material_id}",
        type="REDUCAO_CONSUMO",
        severity=severity,
        title=f"Redução relevante de consumo — {material}",
        description=(
            f"{material} apresentou redução de "
            f"{abs(variation):.2f}% no consumo do último mês, "
            f"passando de {previous:.2f} para {latest:.2f}."
        ),
        entity=material_id,
        rule_origin="rule_consumption_drop",
        evidences=evidences,
        metrics={
            "consumo_mes_anterior": previous,
            "consumo_ultimo_mes": latest,
            "variacao_consumo_pct": variation,
            "limite_variacao_pct": threshold_pct,
        },
        recommended_action=(
            "Verificar se a redução representa mudança operacional, "
            "sazonalidade, alteração de demanda ou possível anomalia "
            "nos registros de consumo."
        ),
    )


def evaluate_consumption_rules(
    features: pd.DataFrame,
) -> tuple[Signal, ...]:
    """
    Executa as regras determinísticas de comportamento de consumo.

    A ordenação final segue o padrão do motor:
    severidade decrescente e ID crescente.
    """

    _validate_columns(features)

    signals: list[Signal] = []

    for _, row in features.iterrows():
        rules = (
            rule_consumption_growth,
            rule_consumption_drop,
        )

        for rule in rules:
            signal = rule(row)

            if signal is not None:
                signals.append(signal)

    return tuple(
        sorted(
            signals,
            key=lambda signal: (
                -signal.priority_rank,
                signal.id,
            ),
        )
    )