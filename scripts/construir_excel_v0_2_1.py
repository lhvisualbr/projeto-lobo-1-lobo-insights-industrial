# -*- coding: utf-8 -*-
"""Gera o XLSX ATUAL com o nome obrigatório da release V0.2.1.

Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.

A V0.2.1 é uma microversão de empacotamento: NENHUMA lógica analítica mudou. Por isso este módulo é apenas um
invólucro: usa `construir_v0_2()` (scripts/construir_excel_v0_2.py, inalterado desde a V0.2 auditada), grava o
resultado em `excel/Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx` e usa o mesmo recálculo/cache da V0.1.1.

O XLSX entregue é byte a byte igual ao XLSX auditado da V0.2 (`excel/lobo_insights_industrial_v0_2.xlsx`, mantido
com o nome histórico). Reconstruir o arquivo com este script produz o mesmo conteúdo (fórmulas e valores), mas não
os mesmos bytes: os metadados do formato variam a cada geração.
"""
import shutil
import tempfile
from pathlib import Path

from construir_excel import LibreOfficeIndisponivel, injetar_cache, recalcular, sha256
from construir_excel_v0_2 import construir_v0_2
from lobo_common import DATA, EXCEL
from lobo_release import NOME_XLSX
from validar_dados import carregar


def gerar(destino):
    """Constrói, recalcula (LibreOffice) e grava o XLSX em `destino`. Levanta SystemExit se houver erro de fórmula."""
    dfs = carregar()
    meta = {"sha": {p.name: sha256(p) for p in sorted(DATA.glob("*.csv"))}}
    tmp = Path(tempfile.mkdtemp(prefix="lobo_build_v021_"))
    try:
        cru = tmp / "cru.xlsx"
        construir_v0_2(dfs, cru, meta)
        try:
            valores, info = recalcular(cru)
        except LibreOfficeIndisponivel as e:
            raise SystemExit("ERRO: " + str(e))
        if info.get("status") != "success" or info.get("total_errors"):
            raise SystemExit("Erros de fórmula após o recálculo – corrija antes de continuar.")
        Path(destino).parent.mkdir(parents=True, exist_ok=True)
        injetar_cache(cru, valores, destino)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return info


def main():
    destino = EXCEL / NOME_XLSX
    info = gerar(destino)
    print("recalc:", info)
    print("gerado:", destino, "sha256:", sha256(destino))


if __name__ == "__main__":
    main()
