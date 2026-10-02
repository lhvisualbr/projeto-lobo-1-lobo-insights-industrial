# -*- coding: utf-8 -*-
"""Validações de qualidade dos CSVs (executadas de verdade; nada é declarado sem rodar).

Uso:  python validar_dados.py
Saída: tabela  Teste | Regra | Resultado | Quantidade de erros | Status
Código de saída 0 = todos PASS; 1 = há FAIL.

Os IDs (Q01..Q36) são os mesmos da aba Qualidade do Excel, permitindo comparar
o resultado das fórmulas do Excel com o resultado independente em Python.
"""
import sys

import pandas as pd

from lobo_common import (CATEGORIAS, CENTROS, CHECKS, CRITICIDADES, DATA, ENC,
                         PERIODO_FIM, PERIODO_INICIO, SEP, TIPOS_MOV, UNIDADES)

ARQUIVOS = {"mat": "materiais_ficticios.csv", "est": "estoque_ficticio.csv",
            "con": "consumo_ficticio.csv"}


def carregar(pasta=DATA):
    """Lê tudo como TEXTO (dtype=str) para que valores estranhos não sejam 'corrigidos' pelo pandas."""
    return {k: pd.read_csv(pasta / v, sep=SEP, encoding=ENC, dtype=str,
                           keep_default_na=False) for k, v in ARQUIVOS.items()}


def _n(serie):
    return pd.to_numeric(serie.str.strip().replace("", pd.NA), errors="coerce")


def _dup(df, col):
    return int(df[col].duplicated(keep=False).sum())


def _nulos(df):
    return int((df.apply(lambda s: s.str.strip() == "")).sum().sum())


def _num(serie, regra):
    """Se houver não-numéricos, conta só eles (espelha a fórmula do Excel); senão aplica a regra."""
    x = _n(serie)
    nao_num = int(x.isna().sum())
    return nao_num if nao_num else int(regra(x).sum())


def _miss(df, mat):
    return int((~df["Material_ID"].isin(mat["Material_ID"])).sum())


def _mis(df, mat, col, numerico=False):
    j = df.merge(mat[["Material_ID", col]], on="Material_ID", how="inner",
                 suffixes=("", "_cat"))
    if numerico:
        return int(((_n(j[col]) - _n(j[col + "_cat"])).abs().fillna(1) > 1e-9).sum())
    return int((j[col] != j[col + "_cat"]).sum())


def _dom(serie, lista):
    return int((~serie.isin(lista)).sum())


def executar(dfs):
    mat, est, con = dfs["mat"], dfs["est"], dfs["con"]
    ini, fim = pd.Timestamp(PERIODO_INICIO), pd.Timestamp(PERIODO_FIM)
    datas = pd.to_datetime(con["Data"], format="%Y-%m-%d", errors="coerce")
    esperados = pd.period_range(ini, fim, freq="M").astype(str)
    meses_presentes = set(datas.dropna().dt.strftime("%Y-%m"))
    r = {
        "Q01": _dup(mat, "Material_ID"), "Q02": _dup(est, "Material_ID"),
        "Q03": _dup(con, "Movimento_ID"),
        "Q04": _nulos(mat), "Q05": _nulos(est), "Q06": _nulos(con),
        "Q07": _num(con["Quantidade"], lambda x: (x <= 0) | (x != x.round())),
        "Q08": _num(mat["Custo_Unitario"], lambda x: x <= 0),
        "Q09": _num(est["Custo_Unitario"], lambda x: x <= 0),
        "Q10": _num(con["Custo_Unitario"], lambda x: x <= 0),
        "Q11": _miss(est, mat), "Q12": _miss(con, mat),
        "Q13": int((~mat["Material_ID"].isin(est["Material_ID"])).sum()),
        "Q14": _mis(est, mat, "Material"), "Q15": _mis(con, mat, "Material"),
        "Q16": _mis(est, mat, "Categoria"), "Q17": _mis(con, mat, "Categoria"),
        "Q18": _mis(est, mat, "Unidade"), "Q19": _mis(con, mat, "Unidade"),
        "Q20": _mis(est, mat, "Custo_Unitario", True), "Q21": _mis(con, mat, "Custo_Unitario", True),
        "Q22": _mis(est, mat, "Criticidade"), "Q23": _mis(est, mat, "Fornecedor"),
        "Q24": int(datas.isna().sum() + ((datas < ini) | (datas > fim)).sum()),
        "Q25": int(sum(m not in meses_presentes for m in esperados)),
        "Q26": _num(est["Estoque_Atual"], lambda x: x < 0),
        "Q27": _num(est["Estoque_Minimo"], lambda x: x < 0),
        "Q28": _num(est["Lead_Time_Dias"], lambda x: (x <= 0) | (x != x.round())),
        "Q29": _dom(mat["Categoria"], CATEGORIAS), "Q30": _dom(est["Categoria"], CATEGORIAS),
        "Q31": _dom(con["Categoria"], CATEGORIAS),
        "Q32": _dom(mat["Unidade"], UNIDADES), "Q33": _dom(mat["Criticidade"], CRITICIDADES),
        "Q34": _dom(con["Centro_Trabalho"], CENTROS), "Q35": _dom(con["Tipo_Movimentacao"], TIPOS_MOV),
        "Q36": int((~mat["Fornecedor"].str.startswith("Fornecedor ")).sum()),
    }
    return r


def tabela(resultados):
    linhas = []
    for cid, teste, regra, _f in CHECKS:
        n = resultados[cid]
        linhas.append({"ID": cid, "Teste": teste, "Regra": regra,
                       "Resultado": "Nenhum erro encontrado" if n == 0 else "Erros encontrados",
                       "Quantidade de erros": n, "Status": "PASS" if n == 0 else "FAIL"})
    return pd.DataFrame(linhas)


def main():
    t = tabela(executar(carregar()))
    with pd.option_context("display.width", 250, "display.max_colwidth", 70):
        print(t.to_string(index=False))
    falhas = int((t["Status"] == "FAIL").sum())
    print(f"\n{len(t)} verificações executadas; {len(t) - falhas} PASS; {falhas} FAIL")
    sys.exit(1 if falhas else 0)


if __name__ == "__main__":
    main()
