# -*- coding: utf-8 -*-
"""
Suíte de validação da V0.3.0 — Lobo Insights Industrial.

Objetivos:
1. Preservar arquivos técnicos históricos da baseline V0.2.1.
2. Executar as 36 verificações de qualidade dos dados.
3. Executar todos os testes novos da V0.3.0 em tests/.

Códigos de saída:
    0 = tudo executado e PASS
    1 = existe FAIL/ERROR
    2 = nenhum FAIL, mas existe SKIP
"""

from __future__ import annotations

import hashlib
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
TESTS = ROOT / "tests"

if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


import validar_dados  # noqa: E402


BASELINE_V021 = {
    "scripts/executar_testes.py":
        "f4bfd733cf271a99f38e734a07786805c3ddf0a4fa1a402b080a77434bfa71e7",

    "scripts/executar_testes_v0_2.py":
        "f502495502ef3d9da6aab457a2c49769e8e5723ad0808bb18e219e13f559969b",

    "scripts/executar_testes_v0_2_1.py":
        "d2bf1501e9ace98ccd94a0d79a49bac9b298b386253bf000c237a2db030201ab",

    "scripts/construir_excel.py":
        "aa02be2f3e53e6f3a6cc826c750787ee2e7f6d4f93ae1d6480977006e30ecdb0",

    "scripts/construir_excel_v0_2.py":
        "e0458a9131a3b5e17b6a5d0ba1ab2e8fb5c077d81dea52b6257e03313c62c596",

    "scripts/construir_excel_v0_2_1.py":
        "dea3cbcd0d3a0e5e74fb6dee79cca1fc813a62fe877e05beef887557df469a6d",

    "scripts/lobo_release.py":
        "91d8e728c5947454a759e9b7c3438c0e2989f0cc11f2f1e02d16ba03f906ec8f",
}


def sha256(caminho: Path) -> str:
    h = hashlib.sha256()

    with caminho.open("rb") as arquivo:
        for bloco in iter(lambda: arquivo.read(1024 * 1024), b""):
            h.update(bloco)

    return h.hexdigest()


def verificar_baseline_historica() -> bool:
    print()
    print("=" * 72)
    print("1. BASELINE HISTORICA V0.2.1")
    print("=" * 72)

    falhas = []

    for relativo, esperado in BASELINE_V021.items():
        caminho = ROOT / relativo

        if not caminho.is_file():
            print(f"FAIL  {relativo} — arquivo ausente")
            falhas.append(relativo)
            continue

        obtido = sha256(caminho)

        if obtido != esperado:
            print(f"FAIL  {relativo} — SHA-256 divergente")
            falhas.append(relativo)
            continue

        print(f"PASS  {relativo}")

    print()
    print(
        f"Baseline histórica: "
        f"{len(BASELINE_V021) - len(falhas)}/{len(BASELINE_V021)} PASS"
    )

    return not falhas


def verificar_qualidade_dados() -> bool:
    print()
    print("=" * 72)
    print("2. QUALIDADE DOS DADOS")
    print("=" * 72)

    dfs = validar_dados.carregar()
    resultados = validar_dados.executar(dfs)
    tabela = validar_dados.tabela(resultados)

    falhas = tabela[tabela["Status"] == "FAIL"]

    for _, linha in tabela.iterrows():
        print(
            f"{linha['ID']}  "
            f"{linha['Status']:<4}  "
            f"{linha['Teste']}  "
            f"(erros={linha['Quantidade de erros']})"
        )

    total = len(tabela)
    qtd_falhas = len(falhas)

    print()
    print(
        f"Qualidade dos dados: "
        f"{total - qtd_falhas}/{total} PASS; "
        f"{qtd_falhas} FAIL"
    )

    return qtd_falhas == 0


def executar_testes_v03():
    print()
    print("=" * 72)
    print("3. TESTES V0.3.0")
    print("=" * 72)

    if not TESTS.is_dir():
        raise RuntimeError(
            f"Pasta de testes não encontrada: {TESTS}"
        )

    loader = unittest.TestLoader()

    suite = loader.discover(
        start_dir=str(TESTS),
        pattern="test_*.py",
        top_level_dir=str(ROOT),
    )

    runner = unittest.TextTestRunner(
        verbosity=2,
    )

    return runner.run(suite)


def main() -> int:
    print("=" * 72)
    print("LOBO INSIGHTS INDUSTRIAL — SUITE V0.3.0")
    print("=" * 72)

    baseline_ok = verificar_baseline_historica()
    dados_ok = verificar_qualidade_dados()

    resultado = executar_testes_v03()

    falhas_unitarias = (
        len(resultado.failures)
        + len(resultado.errors)
    )

    skips = len(resultado.skipped)

    passes_unitarios = (
        resultado.testsRun
        - falhas_unitarias
        - skips
    )

    print()
    print("=" * 72)
    print("RESUMO V0.3.0")
    print("=" * 72)

    print(
        "Baseline histórica:",
        "PASS" if baseline_ok else "FAIL",
    )

    print(
        "Qualidade dos dados:",
        "PASS" if dados_ok else "FAIL",
        "(36 verificações)",
    )

    print(
        "Testes V0.3.0:",
        f"{resultado.testsRun} executados | "
        f"PASS {passes_unitarios} | "
        f"FAIL/ERROR {falhas_unitarias} | "
        f"SKIP {skips}",
    )

    if (
        not baseline_ok
        or not dados_ok
        or falhas_unitarias
    ):
        print()
        print("RESULTADO FINAL: FAIL")
        return 1

    if skips:
        print()
        print("RESULTADO FINAL: PASS COM SKIP")
        return 2

    print()
    print("RESULTADO FINAL: PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())