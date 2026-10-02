from __future__ import annotations

import math
import unicodedata
from typing import Any, Mapping

import pandas as pd

from intelligence.models import (
    Evidence,
    Severity,
    Signal,
)


REQUIRED_COLUMNS = {
    "Material_ID",
    "Material",
    "Criticidade",
    "Estoque_Atual",
    "Estoque_Minimo",
    "Lead_Time_Dias",
    "gap_estoque_minimo",
    "razao_estoque_minimo",
    "cobertura_meses",
    "media_mensal_consumo",
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


def _is_high_criticality(value: object) -> bool:
    return _normalize_text(value) == "alta"


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


def _validate_columns(df: pd.DataFrame) -> None:
    missing = sorted(
        REQUIRED_COLUMNS - set(df.columns)
    )

    if missing:
        raise ValueError(
            "features: colunas obrigatórias ausentes: "
            + ", ".join(missing)
        )


def rule_stock_below_minimum(
    row: Mapping[str, Any],
) -> Signal | None:
    """
    Detecta estoque estritamente abaixo do mínimo.

    Severidade:
    - CRITICO quando a criticidade original é Alta;
    - ALTO nos demais casos.
    """

    atual = _as_float(
        row["Estoque_Atual"]
    )

    minimo = _as_float(
        row["Estoque_Minimo"]
    )

    if atual is None or minimo is None:
        return None

    if atual >= minimo:
        return None

    material_id = str(
        row["Material_ID"]
    )

    material = str(
        row["Material"]
    )

    criticidade = str(
        row["Criticidade"]
    )

    gap = atual - minimo

    severity = (
        Severity.CRITICO
        if _is_high_criticality(criticidade)
        else Severity.ALTO
    )

    return Signal(
        id=f"STOCK-BELOW-MIN-{material_id}",
        type="ESTOQUE_ABAIXO_MINIMO",
        severity=severity,
        title=f"Estoque abaixo do mínimo — {material}",
        description=(
            f"{material} possui estoque atual de "
            f"{atual:g}, abaixo do mínimo de "
            f"{minimo:g}."
        ),
        entity=material_id,
        rule_origin="rule_stock_below_minimum",
        evidences=(
            Evidence(
                source="estoque",
                metric="Estoque_Atual",
                value=atual,
                description=(
                    "Quantidade disponível no estoque."
                ),
            ),
            Evidence(
                source="estoque",
                metric="Estoque_Minimo",
                value=minimo,
                description=(
                    "Nível mínimo registrado para o material."
                ),
            ),
            Evidence(
                source="estoque",
                metric="Criticidade",
                value=criticidade,
                description=(
                    "Criticidade original cadastrada."
                ),
            ),
        ),
        metrics={
            "estoque_atual": atual,
            "estoque_minimo": minimo,
            "gap_estoque_minimo": gap,
        },
        recommended_action=(
            "Avaliar reposição do material."
        ),
    )


def rule_stock_near_minimum(
    row: Mapping[str, Any],
    limit_ratio: float = 1.25,
) -> Signal | None:
    """
    Detecta estoque acima ou igual ao mínimo,
    mas dentro de uma faixa de atenção.

    Regra padrão:
        1.00 <= estoque/minimo <= 1.25

    Itens já abaixo do mínimo não entram aqui.
    """

    ratio = _as_float(
        row["razao_estoque_minimo"]
    )

    atual = _as_float(
        row["Estoque_Atual"]
    )

    minimo = _as_float(
        row["Estoque_Minimo"]
    )

    if (
        ratio is None
        or atual is None
        or minimo is None
    ):
        return None

    if ratio < 1.0:
        return None

    if ratio > limit_ratio:
        return None

    material_id = str(
        row["Material_ID"]
    )

    material = str(
        row["Material"]
    )

    criticidade = str(
        row["Criticidade"]
    )

    severity = (
        Severity.ALTO
        if _is_high_criticality(criticidade)
        else Severity.ATENCAO
    )

    return Signal(
        id=f"STOCK-NEAR-MIN-{material_id}",
        type="ESTOQUE_PROXIMO_MINIMO",
        severity=severity,
        title=f"Estoque próximo do mínimo — {material}",
        description=(
            f"{material} possui estoque atual de "
            f"{atual:g} para mínimo de {minimo:g}, "
            f"razão de {ratio:.2f}."
        ),
        entity=material_id,
        rule_origin="rule_stock_near_minimum",
        evidences=(
            Evidence(
                source="estoque",
                metric="Estoque_Atual",
                value=atual,
            ),
            Evidence(
                source="estoque",
                metric="Estoque_Minimo",
                value=minimo,
            ),
            Evidence(
                source="features",
                metric="razao_estoque_minimo",
                value=ratio,
                description=(
                    "Relação entre estoque atual e mínimo."
                ),
            ),
        ),
        metrics={
            "estoque_atual": atual,
            "estoque_minimo": minimo,
            "razao_estoque_minimo": ratio,
        },
        recommended_action=(
            "Acompanhar o estoque e avaliar necessidade "
            "de reposição."
        ),
    )


def rule_low_coverage(
    row: Mapping[str, Any],
    limit_months: float = 1.0,
) -> Signal | None:
    """
    Detecta cobertura inferior ao limite informado.

    Cobertura:
        estoque atual / média mensal de consumo

    Por padrão, sinaliza cobertura inferior a 1 mês.
    """

    coverage = _as_float(
        row["cobertura_meses"]
    )

    monthly_consumption = _as_float(
        row["media_mensal_consumo"]
    )

    lead_time = _as_float(
        row["Lead_Time_Dias"]
    )

    if coverage is None:
        return None

    if monthly_consumption is None:
        return None

    if monthly_consumption <= 0:
        return None

    if coverage >= limit_months:
        return None

    material_id = str(
        row["Material_ID"]
    )

    material = str(
        row["Material"]
    )

    criticidade = str(
        row["Criticidade"]
    )

    severity = (
        Severity.CRITICO
        if _is_high_criticality(criticidade)
        else Severity.ALTO
    )

    evidences = [
        Evidence(
            source="features",
            metric="cobertura_meses",
            value=coverage,
            description=(
                "Meses estimados de cobertura com base "
                "na média mensal de consumo."
            ),
        ),
        Evidence(
            source="features",
            metric="media_mensal_consumo",
            value=monthly_consumption,
        ),
    ]

    if lead_time is not None:
        evidences.append(
            Evidence(
                source="estoque",
                metric="Lead_Time_Dias",
                value=lead_time,
                description=(
                    "Tempo de reposição cadastrado."
                ),
            )
        )

    return Signal(
        id=f"LOW-COVERAGE-{material_id}",
        type="BAIXA_COBERTURA",
        severity=severity,
        title=f"Baixa cobertura de estoque — {material}",
        description=(
            f"{material} possui cobertura estimada de "
            f"{coverage:.2f} mês(es), abaixo do limite "
            f"de {limit_months:.2f}."
        ),
        entity=material_id,
        rule_origin="rule_low_coverage",
        evidences=tuple(evidences),
        metrics={
            "cobertura_meses": coverage,
            "media_mensal_consumo": monthly_consumption,
            "lead_time_dias": lead_time,
        },
        recommended_action=(
            "Avaliar cobertura e planejamento de reposição."
        ),
    )


def evaluate_inventory_rules(
    features: pd.DataFrame,
) -> tuple[Signal, ...]:
    """
    Executa as regras iniciais de estoque.

    A ordem final é determinística:
    severidade decrescente e ID crescente.
    """

    _validate_columns(
        features
    )

    signals: list[Signal] = []

    for _, row in features.iterrows():
        rules = (
            rule_stock_below_minimum,
            rule_stock_near_minimum,
            rule_low_coverage,
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