from __future__ import annotations

import math
import unicodedata

import pandas as pd

from .models import (
    Evidence,
    Severity,
    Signal,
)


DAYS_PER_MONTH = 30.0

CRITICAL_FACTOR_THRESHOLD = 1.0
HIGH_FACTOR_THRESHOLD = 1.5
ATTENTION_FACTOR_THRESHOLD = 2.0


REQUIRED_COLUMNS = {
    "Material_ID",
    "Material",
    "Lead_Time_Dias",
    "Criticidade",
    "cobertura_meses",
}


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
            "para regras de lead time: "
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


def _normalize_text(
    value: object,
) -> str:
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


def _base_severity(
    replenishment_factor: float,
) -> Severity | None:
    if replenishment_factor < CRITICAL_FACTOR_THRESHOLD:
        return Severity.CRITICO

    if replenishment_factor < HIGH_FACTOR_THRESHOLD:
        return Severity.ALTO

    if replenishment_factor < ATTENTION_FACTOR_THRESHOLD:
        return Severity.ATENCAO

    return None


def _apply_criticality(
    severity: Severity,
    criticality: object,
) -> Severity:
    """
    Materiais de criticidade alta recebem elevação
    de um nível de severidade.

    A criticidade não reduz severidade e nunca
    ultrapassa CRITICO.
    """

    normalized = _normalize_text(
        criticality
    )

    if normalized != "alta":
        return severity

    escalation = {
        Severity.INFO: Severity.ATENCAO,
        Severity.ATENCAO: Severity.ALTO,
        Severity.ALTO: Severity.CRITICO,
        Severity.CRITICO: Severity.CRITICO,
    }

    return escalation[
        severity
    ]


def rule_replenishment_risk(
    row: pd.Series,
) -> Signal | None:
    """
    Identifica risco de reposição comparando a cobertura
    disponível com o lead time cadastrado.

    fator_reposicao =
        cobertura_em_dias / lead_time_em_dias

    Interpretação padrão da V0.3.0:
    - fator < 1.0: CRITICO;
    - fator < 1.5: ALTO;
    - fator < 2.0: ATENCAO;
    - fator >= 2.0: sem sinal.

    Materiais de criticidade Alta recebem elevação
    de um nível de severidade.
    """

    material_id = str(
        row["Material_ID"]
    ).strip()

    material = str(
        row["Material"]
    ).strip()

    criticality = str(
        row["Criticidade"]
    ).strip()

    lead_time = _safe_float(
        row["Lead_Time_Dias"]
    )

    coverage_months = _safe_float(
        row["cobertura_meses"]
    )

    if (
        lead_time is None
        or coverage_months is None
    ):
        return None

    if lead_time <= 0:
        return None

    if coverage_months < 0:
        return None

    coverage_days = (
        coverage_months
        * DAYS_PER_MONTH
    )

    replenishment_factor = (
        coverage_days
        / lead_time
    )

    severity = _base_severity(
        replenishment_factor
    )

    if severity is None:
        return None

    base_severity = severity

    severity = _apply_criticality(
        severity,
        criticality,
    )

    evidences = (
        Evidence(
            source="estoque",
            metric="Lead_Time_Dias",
            value=lead_time,
            description=(
                "Tempo de reposição cadastrado "
                "para o material."
            ),
        ),
        Evidence(
            source="estoque",
            metric="cobertura_meses",
            value=coverage_months,
            description=(
                "Cobertura estimada de estoque "
                "em meses."
            ),
        ),
        Evidence(
            source="estoque",
            metric="Criticidade",
            value=criticality,
            description=(
                "Criticidade operacional "
                "cadastrada para o material."
            ),
        ),
    )

    escalation_applied = (
        severity != base_severity
    )

    description = (
        f"{material} possui cobertura estimada "
        f"de {coverage_days:.2f} dia(s) para um "
        f"lead time de {lead_time:.2f} dia(s), "
        f"resultando em fator de reposição "
        f"{replenishment_factor:.2f}."
    )

    if escalation_applied:
        description += (
            " A severidade foi elevada devido "
            "à criticidade Alta do material."
        )

    return Signal(
        id=f"REPLENISHMENT-RISK-{material_id}",
        type="RISCO_REPOSICAO",
        severity=severity,
        title=(
            f"Risco de reposição — "
            f"{material}"
        ),
        description=description,
        entity=material_id,
        rule_origin="rule_replenishment_risk",
        evidences=evidences,
        metrics={
            "lead_time_dias": lead_time,
            "cobertura_meses": coverage_months,
            "cobertura_dias": coverage_days,
            "fator_reposicao": replenishment_factor,
            "criticidade_alta": (
                _normalize_text(
                    criticality
                )
                == "alta"
            ),
            "severidade_elevada": (
                escalation_applied
            ),
            "limite_critico": (
                CRITICAL_FACTOR_THRESHOLD
            ),
            "limite_alto": (
                HIGH_FACTOR_THRESHOLD
            ),
            "limite_atencao": (
                ATTENTION_FACTOR_THRESHOLD
            ),
        },
        recommended_action=(
            "Revisar o planejamento de reposição, "
            "o ponto de pedido e a cobertura disponível "
            "considerando o lead time e a criticidade "
            "operacional do material."
        ),
    )


def evaluate_lead_time_rules(
    features: pd.DataFrame,
) -> tuple[Signal, ...]:
    """
    Executa as regras determinísticas de risco
    de reposição.

    A ordenação segue o padrão do projeto:
    severidade decrescente e ID crescente.
    """

    _validate_columns(
        features
    )

    signals: list[Signal] = []

    for _, row in features.iterrows():
        signal = rule_replenishment_risk(
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