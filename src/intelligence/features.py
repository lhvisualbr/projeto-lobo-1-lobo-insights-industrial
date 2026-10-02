from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

import pandas as pd


ESTOQUE_REQUIRED = {
    "Material_ID",
    "Material",
    "Categoria",
    "Unidade",
    "Estoque_Atual",
    "Estoque_Minimo",
    "Custo_Unitario",
    "Lead_Time_Dias",
    "Fornecedor",
    "Criticidade",
}

CONSUMO_REQUIRED = {
    "Movimento_ID",
    "Data",
    "Material_ID",
    "Material",
    "Categoria",
    "Quantidade",
    "Unidade",
    "Custo_Unitario",
    "Centro_Trabalho",
    "Tipo_Movimentacao",
}


@dataclass(frozen=True, slots=True)
class FeatureContext:
    """
    Metadados objetivos do período analisado.
    """

    data_inicio: str | None
    data_fim: str | None
    meses_analisados: int
    total_movimentacoes: int
    total_quantidade: float
    total_custo: float

    def to_dict(self) -> dict[str, object]:
        return {
            "data_inicio": self.data_inicio,
            "data_fim": self.data_fim,
            "meses_analisados": self.meses_analisados,
            "total_movimentacoes": self.total_movimentacoes,
            "total_quantidade": self.total_quantidade,
            "total_custo": self.total_custo,
        }


def _validate_columns(
    df: pd.DataFrame,
    required: Iterable[str],
    dataset_name: str,
) -> None:
    missing = sorted(set(required) - set(df.columns))

    if missing:
        raise ValueError(
            f"{dataset_name}: colunas obrigatórias ausentes: "
            + ", ".join(missing)
        )


def _safe_divide(
    numerator: pd.Series,
    denominator: pd.Series,
) -> pd.Series:
    denominator = denominator.astype(float)

    return numerator.astype(float).div(
        denominator.where(denominator != 0)
    )


def build_stock_features(
    estoque: pd.DataFrame,
) -> pd.DataFrame:
    """
    Calcula métricas objetivas de estoque.

    Nenhuma severidade ou recomendação é produzida aqui.
    """

    _validate_columns(
        estoque,
        ESTOQUE_REQUIRED,
        "estoque",
    )

    result = estoque.copy()

    result["gap_estoque_minimo"] = (
        result["Estoque_Atual"]
        - result["Estoque_Minimo"]
    )

    result["razao_estoque_minimo"] = _safe_divide(
        result["Estoque_Atual"],
        result["Estoque_Minimo"],
    )

    result["percentual_gap_minimo"] = (
        _safe_divide(
            result["gap_estoque_minimo"],
            result["Estoque_Minimo"],
        )
        * 100
    )

    result["valor_estoque_atual"] = (
        result["Estoque_Atual"]
        * result["Custo_Unitario"]
    )

    result["valor_estoque_minimo"] = (
        result["Estoque_Minimo"]
        * result["Custo_Unitario"]
    )

    return result


def build_consumption_features(
    consumo: pd.DataFrame,
) -> tuple[pd.DataFrame, FeatureContext]:
    """
    Agrega consumo e custo por material.

    Calcula:
    - quantidade total consumida;
    - custo total;
    - quantidade de movimentações;
    - média mensal;
    - participação no consumo;
    - participação no custo;
    - consumo do último mês;
    - consumo do mês anterior;
    - variação percentual entre os dois últimos meses.
    """

    _validate_columns(
        consumo,
        CONSUMO_REQUIRED,
        "consumo",
    )

    df = consumo.copy()

    df["Data"] = pd.to_datetime(
        df["Data"],
        errors="raise",
    )

    df["custo_movimentacao"] = (
        df["Quantidade"]
        * df["Custo_Unitario"]
    )

    df["mes"] = df["Data"].dt.to_period("M")

    meses = sorted(df["mes"].dropna().unique())

    total_quantidade = float(
        df["Quantidade"].sum()
    )

    total_custo = float(
        df["custo_movimentacao"].sum()
    )

    agrupado = (
        df.groupby(
            "Material_ID",
            as_index=False,
        )
        .agg(
            quantidade_consumida=(
                "Quantidade",
                "sum",
            ),
            custo_consumo=(
                "custo_movimentacao",
                "sum",
            ),
            movimentacoes=(
                "Movimento_ID",
                "count",
            ),
        )
    )

    meses_analisados = len(meses)

    if meses_analisados:
        agrupado["media_mensal_consumo"] = (
            agrupado["quantidade_consumida"]
            / meses_analisados
        )
    else:
        agrupado["media_mensal_consumo"] = 0.0

    if total_quantidade:
        agrupado["participacao_consumo_pct"] = (
            agrupado["quantidade_consumida"]
            / total_quantidade
            * 100
        )
    else:
        agrupado["participacao_consumo_pct"] = 0.0

    if total_custo:
        agrupado["participacao_custo_pct"] = (
            agrupado["custo_consumo"]
            / total_custo
            * 100
        )
    else:
        agrupado["participacao_custo_pct"] = 0.0

    mensal = (
        df.groupby(
            ["Material_ID", "mes"],
            as_index=False,
        )["Quantidade"]
        .sum()
    )

    if meses:
        ultimo_mes = meses[-1]

        ultimo = (
            mensal[mensal["mes"] == ultimo_mes]
            .set_index("Material_ID")["Quantidade"]
        )

        agrupado["consumo_ultimo_mes"] = (
            agrupado["Material_ID"]
            .map(ultimo)
            .fillna(0)
            .astype(float)
        )
    else:
        agrupado["consumo_ultimo_mes"] = 0.0

    if len(meses) >= 2:
        mes_anterior = meses[-2]

        anterior = (
            mensal[mensal["mes"] == mes_anterior]
            .set_index("Material_ID")["Quantidade"]
        )

        agrupado["consumo_mes_anterior"] = (
            agrupado["Material_ID"]
            .map(anterior)
            .fillna(0)
            .astype(float)
        )
    else:
        agrupado["consumo_mes_anterior"] = 0.0

    anterior = agrupado["consumo_mes_anterior"]
    ultimo = agrupado["consumo_ultimo_mes"]

    agrupado["variacao_consumo_pct"] = (
        (ultimo - anterior)
        .div(anterior.where(anterior != 0))
        * 100
    )

    context = FeatureContext(
        data_inicio=(
            df["Data"].min().date().isoformat()
            if not df.empty
            else None
        ),
        data_fim=(
            df["Data"].max().date().isoformat()
            if not df.empty
            else None
        ),
        meses_analisados=meses_analisados,
        total_movimentacoes=len(df),
        total_quantidade=total_quantidade,
        total_custo=total_custo,
    )

    return agrupado, context


def build_material_features(
    estoque: pd.DataFrame,
    consumo: pd.DataFrame,
) -> tuple[pd.DataFrame, FeatureContext]:
    """
    Constrói a tabela analítica principal por material.

    Combina dados objetivos de estoque e consumo.
    """

    stock = build_stock_features(
        estoque
    )

    consumption, context = (
        build_consumption_features(
            consumo
        )
    )

    result = stock.merge(
        consumption,
        on="Material_ID",
        how="left",
        validate="one_to_one",
    )

    consumption_columns = [
        "quantidade_consumida",
        "custo_consumo",
        "movimentacoes",
        "media_mensal_consumo",
        "participacao_consumo_pct",
        "participacao_custo_pct",
        "consumo_ultimo_mes",
        "consumo_mes_anterior",
    ]

    for column in consumption_columns:
        result[column] = (
            result[column]
            .fillna(0)
        )

    result["cobertura_meses"] = _safe_divide(
        result["Estoque_Atual"],
        result["media_mensal_consumo"],
    )

    return result, context