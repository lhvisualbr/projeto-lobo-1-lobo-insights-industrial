from __future__ import annotations

import math
import unicodedata

import pandas as pd

from .models import (
    Evidence,
    Severity,
    Signal,
)


SUPPLIER_COST_THRESHOLD_PCT = 25.0
SUPPLIER_COST_HIGH_THRESHOLD_PCT = 40.0

CRITICAL_MATERIAL_THRESHOLD = 2
CRITICAL_MATERIAL_CRITICAL_THRESHOLD = 3


REQUIRED_COLUMNS = {
    "Material_ID",
    "Fornecedor",
    "Criticidade",
    "custo_consumo",
    "participacao_custo_pct",
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
            "para regras de fornecedor: "
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


def _supplier_key(
    supplier: object,
) -> str:
    normalized = _normalize_text(
        supplier
    )

    characters: list[str] = []

    previous_dash = False

    for char in normalized:
        if char.isalnum():
            characters.append(
                char.upper()
            )
            previous_dash = False
        else:
            if (
                characters
                and not previous_dash
            ):
                characters.append("-")
                previous_dash = True

    result = "".join(
        characters
    ).strip("-")

    return result or "SEM-ID"


def build_supplier_features(
    features: pd.DataFrame,
) -> pd.DataFrame:
    """
    Agrega as features de materiais no nível de fornecedor.

    Retorna uma linha por fornecedor com:
    - quantidade de materiais;
    - quantidade de materiais de criticidade Alta;
    - custo total associado;
    - participação no custo total.
    """

    _validate_columns(
        features
    )

    rows: list[dict[str, object]] = []

    for supplier, group in features.groupby(
        "Fornecedor",
        dropna=False,
        sort=True,
    ):
        supplier_name = str(
            supplier
        ).strip()

        if (
            not supplier_name
            or _normalize_text(
                supplier_name
            )
            in {
                "nan",
                "none",
            }
        ):
            continue

        material_ids = (
            group["Material_ID"]
            .astype(str)
            .str.strip()
        )

        critical_mask = (
            group["Criticidade"]
            .map(_normalize_text)
            == "alta"
        )

        critical_material_ids = (
            group.loc[
                critical_mask,
                "Material_ID",
            ]
            .astype(str)
            .str.strip()
        )

        cost_values = pd.to_numeric(
            group["custo_consumo"],
            errors="coerce",
        )

        share_values = pd.to_numeric(
            group["participacao_custo_pct"],
            errors="coerce",
        )

        total_cost = float(
            cost_values.fillna(0.0).sum()
        )

        cost_share = float(
            share_values.fillna(0.0).sum()
        )

        rows.append(
            {
                "Fornecedor": supplier_name,
                "materiais": int(
                    material_ids.nunique()
                ),
                "materiais_criticos": int(
                    critical_material_ids.nunique()
                ),
                "custo_total": total_cost,
                "participacao_custo_pct": (
                    cost_share
                ),
            }
        )

    return pd.DataFrame(
        rows,
        columns=[
            "Fornecedor",
            "materiais",
            "materiais_criticos",
            "custo_total",
            "participacao_custo_pct",
        ],
    )


def rule_supplier_cost_concentration(
    row: pd.Series,
    threshold_pct: float = (
        SUPPLIER_COST_THRESHOLD_PCT
    ),
    high_threshold_pct: float = (
        SUPPLIER_COST_HIGH_THRESHOLD_PCT
    ),
) -> Signal | None:
    """
    Identifica concentração financeira relevante
    em um único fornecedor.

    Limites padrão da V0.3.0:
    - >= 25% do custo total: ATENCAO;
    - >= 40% do custo total: ALTO.
    """

    supplier = str(
        row["Fornecedor"]
    ).strip()

    material_count = _safe_float(
        row["materiais"]
    )

    total_cost = _safe_float(
        row["custo_total"]
    )

    cost_share = _safe_float(
        row["participacao_custo_pct"]
    )

    if (
        material_count is None
        or total_cost is None
        or cost_share is None
    ):
        return None

    if material_count <= 0:
        return None

    if total_cost <= 0:
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
            source="fornecedor",
            metric="participacao_custo_pct",
            value=cost_share,
            description=(
                "Participação do fornecedor "
                "no custo total de consumo."
            ),
        ),
        Evidence(
            source="fornecedor",
            metric="custo_total",
            value=total_cost,
            description=(
                "Custo total associado aos materiais "
                "do fornecedor no período analisado."
            ),
        ),
        Evidence(
            source="fornecedor",
            metric="materiais",
            value=material_count,
            description=(
                "Quantidade de materiais distintos "
                "associados ao fornecedor."
            ),
        ),
    )

    return Signal(
        id=(
            "SUPPLIER-COST-CONCENTRATION-"
            f"{_supplier_key(supplier)}"
        ),
        type="CONCENTRACAO_FORNECEDOR",
        severity=severity,
        title=(
            "Concentração financeira por fornecedor — "
            f"{supplier}"
        ),
        description=(
            f"{supplier} concentra "
            f"{cost_share:.2f}% do custo total "
            f"de consumo e atende "
            f"{int(material_count)} material(is)."
        ),
        entity=supplier,
        rule_origin=(
            "rule_supplier_cost_concentration"
        ),
        evidences=evidences,
        metrics={
            "participacao_custo_pct": (
                cost_share
            ),
            "custo_total": total_cost,
            "materiais": int(
                material_count
            ),
            "limite_concentracao_pct": (
                threshold_pct
            ),
            "limite_alto_pct": (
                high_threshold_pct
            ),
        },
        recommended_action=(
            "Avaliar a dependência financeira deste "
            "fornecedor, alternativas de fornecimento "
            "e estratégias para reduzir exposição "
            "a uma única fonte."
        ),
    )


def rule_supplier_critical_exposure(
    row: pd.Series,
    threshold: int = (
        CRITICAL_MATERIAL_THRESHOLD
    ),
    critical_threshold: int = (
        CRITICAL_MATERIAL_CRITICAL_THRESHOLD
    ),
) -> Signal | None:
    """
    Identifica fornecedores responsáveis por múltiplos
    materiais de criticidade Alta.

    Limites padrão da V0.3.0:
    - >= 2 materiais críticos: ALTO;
    - >= 3 materiais críticos: CRITICO.
    """

    supplier = str(
        row["Fornecedor"]
    ).strip()

    material_count = _safe_float(
        row["materiais"]
    )

    critical_count = _safe_float(
        row["materiais_criticos"]
    )

    cost_share = _safe_float(
        row["participacao_custo_pct"]
    )

    if (
        material_count is None
        or critical_count is None
        or cost_share is None
    ):
        return None

    critical_count_int = int(
        critical_count
    )

    if critical_count_int < threshold:
        return None

    severity = (
        Severity.CRITICO
        if critical_count_int
        >= critical_threshold
        else Severity.ALTO
    )

    evidences = (
        Evidence(
            source="fornecedor",
            metric="materiais_criticos",
            value=critical_count_int,
            description=(
                "Quantidade de materiais distintos "
                "de criticidade Alta atendidos "
                "pelo fornecedor."
            ),
        ),
        Evidence(
            source="fornecedor",
            metric="materiais",
            value=int(
                material_count
            ),
            description=(
                "Quantidade total de materiais "
                "associados ao fornecedor."
            ),
        ),
        Evidence(
            source="fornecedor",
            metric="participacao_custo_pct",
            value=cost_share,
            description=(
                "Participação do fornecedor "
                "no custo total de consumo."
            ),
        ),
    )

    return Signal(
        id=(
            "SUPPLIER-CRITICAL-EXPOSURE-"
            f"{_supplier_key(supplier)}"
        ),
        type="EXPOSICAO_FORNECEDOR_CRITICO",
        severity=severity,
        title=(
            "Exposição de materiais críticos — "
            f"{supplier}"
        ),
        description=(
            f"{supplier} atende "
            f"{critical_count_int} material(is) "
            "de criticidade Alta, entre "
            f"{int(material_count)} material(is) "
            "associados."
        ),
        entity=supplier,
        rule_origin=(
            "rule_supplier_critical_exposure"
        ),
        evidences=evidences,
        metrics={
            "materiais": int(
                material_count
            ),
            "materiais_criticos": (
                critical_count_int
            ),
            "participacao_custo_pct": (
                cost_share
            ),
            "limite_materiais_criticos": (
                threshold
            ),
            "limite_critico": (
                critical_threshold
            ),
        },
        recommended_action=(
            "Avaliar alternativas de fornecimento "
            "para os materiais críticos, contingência "
            "de abastecimento e risco de dependência "
            "do fornecedor."
        ),
    )


def evaluate_supplier_rules(
    features: pd.DataFrame,
) -> tuple[Signal, ...]:
    """
    Executa as regras determinísticas de inteligência
    de fornecedores.

    As features de materiais são agregadas internamente
    para o nível de fornecedor.

    A ordenação segue o padrão do projeto:
    severidade decrescente e ID crescente.
    """

    supplier_features = (
        build_supplier_features(
            features
        )
    )

    signals: list[Signal] = []

    for _, row in supplier_features.iterrows():
        rules = (
            rule_supplier_cost_concentration,
            rule_supplier_critical_exposure,
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