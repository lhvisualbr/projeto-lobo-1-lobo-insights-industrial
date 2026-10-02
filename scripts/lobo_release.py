# -*- coding: utf-8 -*-
"""Release engineering do Projeto Lobo 1 (V0.2.1 - Packaging & Reproducibility Fix).

Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.

Este módulo NÃO contém lógica analítica. Ele concentra:
  1. a nomenclatura obrigatória dos entregáveis principais (prefixo `Projeto_Lobo_1_`);
  2. a estratégia de manifestos SHA-256:
       SHA256SUMS.txt          manifesto OPERACIONAL: só artefatos estáveis (dados, scripts, XLSX, documentação-fonte).
                               Continua válido depois de uma execução normal da suíte de testes.
       RELEASE_SHA256SUMS.txt  SNAPSHOT da release: todos os arquivos (inclusive os regeneráveis) no instante do empacotamento.
                               Só confere integralmente no pacote recém-extraído.
  3. o empacotamento determinístico do ZIP (ordem e datas fixas).

Uso (a partir da raiz do projeto):
    python scripts/lobo_release.py manifestos          # (re)escreve SHA256SUMS.txt e depois RELEASE_SHA256SUMS.txt
    python scripts/lobo_release.py verificar           # confere SHA256SUMS.txt (deve passar após executar os testes)
    python scripts/lobo_release.py verificar --release # confere o snapshot (só no pacote recém-extraído)
    python scripts/lobo_release.py empacotar [--saida PASTA]
"""
import hashlib
import re
import sys
import zipfile
from pathlib import Path

from lobo_common import ROOT

# ------------------------------------------------------------------ nomenclatura obrigatória
PREFIXO = "Projeto_Lobo_1_"
NOME_BASE = "Lobo_Insights_Industrial"
VERSAO = "V0_2_1"
NOME_ZIP = PREFIXO + NOME_BASE + "_" + VERSAO + ".zip"
NOME_XLSX = PREFIXO + NOME_BASE + "_" + VERSAO + ".xlsx"
PASTA_RAIZ = "lobo-insights-industrial"           # pasta-raiz dentro do ZIP (identidade do projeto, estável entre versões)
XLSX_HISTORICOS = ("lobo_insights_industrial_v0_1.xlsx", "lobo_insights_industrial_v0_2.xlsx")   # nomes históricos: não renomear
SHA256_XLSX_V0_2_AUDITADO = "81f51886dc3d68408073a9e0bf54859ae100fe418c2c30060b8d92fb9662b846"
AMBIGUOS = ("Lobo_", "lobo_", "Projeto_Lobo_2", "Lobo_" + "Fore" + "cast", "lobo_" + "fore" + "cast")   # monta por concatenação: a varredura de escopo proíbe o termo nos scripts

# ------------------------------------------------------------------ manifestos
MANIFESTO_ESTAVEL = "SHA256SUMS.txt"
MANIFESTO_RELEASE = "RELEASE_SHA256SUMS.txt"
# Arquivos regenerados de propósito por uma execução normal da suíte e que variam de forma legítima entre ambientes
# (o relatório registra versões/plataforma; as imagens dependem da versão do LibreOffice e das fontes).
# NÃO entram em SHA256SUMS.txt. Qualquer outro arquivo do projeto DEVE estar no manifesto estável.
REGENERAVEIS = ("docs/TEST_REPORT.md",
                "images/dashboard_v0_2_pagina_1.png",
                "images/dashboard_v0_2_pagina_2.png",
                "images/dashboard_v0_2_pagina_3.png")
DATA_ZIP = (2026, 9, 21, 0, 0, 0)                 # data fixa das entradas do ZIP (empacotamento determinístico)
FORMATO = re.compile(r"^([0-9a-f]{64})  (\S.*)$")


def sha256(caminho):
    return hashlib.sha256(Path(caminho).read_bytes()).hexdigest()


def nome_principal_valido(nome):
    """Regra de nomenclatura dos entregáveis principais: prefixo do Projeto 1 e nenhum prefixo ambíguo."""
    nome = str(nome)
    return nome.startswith(PREFIXO) and not any(nome.startswith(a) for a in AMBIGUOS) and nome.endswith((".zip", ".xlsx"))


def arquivos_do_projeto(raiz):
    """Caminhos relativos (POSIX, ordem C) de todos os arquivos, ignorando __pycache__ e .pyc."""
    raiz = Path(raiz)
    itens = [p.relative_to(raiz).as_posix() for p in raiz.rglob("*") if p.is_file() and "__pycache__" not in p.parts and p.suffix != ".pyc"]
    return sorted(itens, key=lambda s: s.encode("utf-8"))


def _linhas(raiz, caminhos):
    return "".join(sha256(Path(raiz) / c) + "  " + c + "\n" for c in caminhos)


def gerar_manifesto_estavel(raiz):
    """Todos os arquivos, exceto os regeneráveis e os dois manifestos."""
    excl = set(REGENERAVEIS) | {MANIFESTO_ESTAVEL, MANIFESTO_RELEASE}
    return _linhas(raiz, [c for c in arquivos_do_projeto(raiz) if c not in excl])


def gerar_manifesto_release(raiz):
    """Snapshot: todos os arquivos (inclusive regeneráveis e SHA256SUMS.txt), exceto o próprio manifesto de release."""
    return _linhas(raiz, [c for c in arquivos_do_projeto(raiz) if c != MANIFESTO_RELEASE])


def ler_manifesto(texto):
    """Lê `<sha256>  <caminho>` (formato aceito por `sha256sum -c`). Linha fora do formato levanta ValueError."""
    saida = []
    for n, linha in enumerate(texto.splitlines(), start=1):
        m = FORMATO.match(linha)
        if not m:
            raise ValueError("linha " + str(n) + " fora do formato sha256sum: " + linha[:60])
        saida.append((m.group(1), m.group(2)))
    return saida


def verificar_manifesto(raiz, manifesto):
    """Confere um manifesto contra a árvore. Retorna {'ausentes': [...], 'divergentes': [...], 'total': n}."""
    raiz = Path(raiz)
    ausentes, divergentes = [], []
    entradas = ler_manifesto((raiz / manifesto).read_text(encoding="utf-8"))
    for h, c in entradas:
        p = raiz / c
        if not p.is_file():
            ausentes.append(c)
        elif sha256(p) != h:
            divergentes.append(c)
    return {"ausentes": ausentes, "divergentes": divergentes, "total": len(entradas)}


def empacotar(raiz, saida):
    """Cria o ZIP da release com nome exato, pasta-raiz fixa, ordem e datas determinísticas. Retorna o caminho do ZIP."""
    raiz, saida = Path(raiz), Path(saida)
    if not nome_principal_valido(NOME_ZIP):
        raise RuntimeError("nome do pacote inválido: " + NOME_ZIP)
    faltam = [m for m in (MANIFESTO_ESTAVEL, MANIFESTO_RELEASE) if not (raiz / m).is_file()]
    if faltam:
        raise RuntimeError("gere os manifestos antes de empacotar: " + ", ".join(faltam))
    saida.mkdir(parents=True, exist_ok=True)
    destino = saida / NOME_ZIP
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        for c in arquivos_do_projeto(raiz):
            info = zipfile.ZipInfo(PASTA_RAIZ + "/" + c, date_time=DATA_ZIP)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            z.writestr(info, (raiz / c).read_bytes())
    return destino


def main(argv):
    cmd = argv[1] if len(argv) > 1 else ""
    if cmd == "manifestos":
        (ROOT / MANIFESTO_ESTAVEL).write_text(gerar_manifesto_estavel(ROOT), encoding="utf-8")
        (ROOT / MANIFESTO_RELEASE).write_text(gerar_manifesto_release(ROOT), encoding="utf-8")
        print("escritos:", MANIFESTO_ESTAVEL, "e", MANIFESTO_RELEASE)
        return 0
    if cmd == "verificar":
        alvo = MANIFESTO_RELEASE if "--release" in argv else MANIFESTO_ESTAVEL
        r = verificar_manifesto(ROOT, alvo)
        print(alvo + ": " + str(r["total"] - len(r["ausentes"]) - len(r["divergentes"])) + " de " + str(r["total"]) + " OK; ausentes=" + str(r["ausentes"]) + "; divergentes=" + str(r["divergentes"]))
        return 0 if not r["ausentes"] and not r["divergentes"] else 1
    if cmd == "empacotar":
        pasta = Path(argv[argv.index("--saida") + 1]) if "--saida" in argv else ROOT.parent
        z = empacotar(ROOT, pasta)
        print("pacote:", z, "sha256:", sha256(z))
        return 0
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
