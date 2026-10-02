# -*- coding: utf-8 -*-
"""Suíte de testes da V0.2.1 (Packaging & Reproducibility Fix). Tudo aqui é EXECUTADO de verdade.

Uso:
    python executar_testes_v0_2_1.py              # roda tudo e (re)escreve docs/TEST_REPORT.md e images/dashboard_v0_2_pagina_N.png
    python executar_testes_v0_2_1.py --verificar  # roda tudo, NÃO escreve relatório nem imagens

Composição (numeração contínua):
    T01-T103  suíte da V0.1.1 (executar_testes.py), executada SEM alteração;
    T104-T174 bateria da V0.2 (executar_testes_v0_2.py), executada SEM alteração e AGORA SOBRE O XLSX COM O NOME DA V0.2.1;
    T175+     testes novos da V0.2.1: nomenclatura, manifestos SHA-256, regressão contra a V0.2 auditada.

Uma execução normal reescreve SOMENTE os arquivos "regeneráveis" (lobo_release.REGENERAVEIS). Eles não fazem parte do
manifesto operacional SHA256SUMS.txt; por isso `sha256sum -c SHA256SUMS.txt` continua passando depois da execução.
Códigos de saída: 0 = tudo executado e PASS; 1 = há FAIL; 2 = nenhum FAIL, mas há SKIP.

Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.
"""
import contextlib
import inspect
import io
import platform
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl
import pandas as pd

import construir_excel_v0_2_1 as gen
import executar_testes as base
import executar_testes_v0_2 as v02
import lobo_release as rel
import validar_dados
from construir_excel import localizar_soffice
from lobo_common import DATA, DOCS, EXCEL, ROOT

t = base.t
AVISO = base.AVISO
NOME_XLSX_ESPERADO = "Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx"      # literais independentes das constantes do módulo
NOME_ZIP_ESPERADO = "Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.zip"
REGENERAVEIS_ESPERADOS = ["docs/TEST_REPORT.md", "images/dashboard_v0_2_pagina_1.png", "images/dashboard_v0_2_pagina_2.png", "images/dashboard_v0_2_pagina_3.png"]
XLSX_ATUAL = EXCEL / NOME_XLSX_ESPERADO
XLSX_V02 = EXCEL / "lobo_insights_industrial_v0_2.xlsx"
HASH_XLSX_V02 = "81f51886dc3d68408073a9e0bf54859ae100fe418c2c30060b8d92fb9662b846"
EXTRA = {"controles": []}


# ================================================================== utilidades
def copiar_arvore(destino):
    """Cópia da árvore do projeto (sem __pycache__) para experimentos destrutivos."""
    shutil.copytree(ROOT, destino, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))


def nao_contabilizados(raiz):
    """Arquivos da árvore que não estão no manifesto estável, não são regeneráveis e não são os manifestos."""
    listados = {c for _, c in rel.ler_manifesto((Path(raiz) / rel.MANIFESTO_ESTAVEL).read_text(encoding="utf-8"))}
    livres = set(rel.REGENERAVEIS) | {rel.MANIFESTO_ESTAVEL, rel.MANIFESTO_RELEASE}
    return sorted(set(rel.arquivos_do_projeto(raiz)) - listados - livres)


def comparar_xlsx(a, b):
    """Compara todas as células (fórmulas/constantes e valores em cache) de dois XLSX. Retorna (n_celulas, dif_formulas, dif_valores, n_graficos_a, n_graficos_b)."""
    fa, fb = openpyxl.load_workbook(a), openpyxl.load_workbook(b)
    va, vb = openpyxl.load_workbook(a, data_only=True), openpyxl.load_workbook(b, data_only=True)
    n = dif_f = dif_v = 0
    for ws in fa.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                n += 1
                dif_f += fb[ws.title][c.coordinate].value != c.value
                x, y = va[ws.title][c.coordinate].value, vb[ws.title][c.coordinate].value
                if isinstance(x, (int, float)) and isinstance(y, (int, float)) and not isinstance(x, bool):
                    dif_v += abs(x - y) > 1e-9
                else:
                    dif_v += (x or None) != (y or None)
    graficos = (sum(len(w._charts) for w in fa.worksheets), sum(len(w._charts) for w in fb.worksheets))
    return n, dif_f, dif_v, graficos


# ================================================================== testes da V0.2.1
def testes_v0_2_1():
    # ---------- nomenclatura
    t("Nomenclatura (V0.2.1)", "o XLSX atual se chama EXATAMENTE Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx e existe em excel/", NOME_XLSX_ESPERADO,
      rel.NOME_XLSX, rel.NOME_XLSX == NOME_XLSX_ESPERADO and XLSX_ATUAL.is_file())
    t("Nomenclatura (V0.2.1)", "o pacote da release se chama EXATAMENTE Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.zip (nome esperado, constante do módulo de release)", NOME_ZIP_ESPERADO, rel.NOME_ZIP,
      rel.NOME_ZIP == NOME_ZIP_ESPERADO)
    t("Nomenclatura (V0.2.1)", "XLSX e ZIP começam com 'Projeto_Lobo_1_' e têm o mesmo radical (só muda a extensão)", True,
      rel.NOME_XLSX.startswith("Projeto_Lobo_1_") and rel.NOME_ZIP.startswith("Projeto_Lobo_1_") and rel.NOME_XLSX[:-5] == rel.NOME_ZIP[:-4])
    lobo_fc = "Lobo_" + "Fore" + "cast_AI_V0_1.zip"
    nega = ["Lobo_Insights_Industrial_V0_2_1.zip", "lobo_insights_industrial_v0_2.xlsx", "Lobo_Insights_Industrial_V0_2.zip", lobo_fc,
            "Projeto_Lobo_2_Lobo_Insights_Industrial_V0_2_1.zip", "lobo_" + "fore" + "cast_v0_1.xlsx", "Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.txt"]
    t("Nomenclatura (V0.2.1)", "a regra de validação aceita os 2 nomes obrigatórios e REJEITA nomes ambíguos (Lobo_…, lobo_…, projeto/produto 2, extensão errada)", "2 aceitos / 0 dos 7 ambíguos aceitos",
      f"{sum(rel.nome_principal_valido(n) for n in (NOME_XLSX_ESPERADO, NOME_ZIP_ESPERADO))} aceitos / {sum(rel.nome_principal_valido(n) for n in nega)} dos {len(nega)} ambíguos aceitos",
      all(rel.nome_principal_valido(n) for n in (NOME_XLSX_ESPERADO, NOME_ZIP_ESPERADO)) and not any(rel.nome_principal_valido(n) for n in nega))
    amb = re.compile(r"^(Lobo_|lobo_|Lobo_" + "Fore" + "cast|Projeto_Lobo_2)")
    t("Nomenclatura (V0.2.1)", "os entregáveis principais desta versão (XLSX atual e ZIP) não casam com nenhum padrão ambíguo", [], [n for n in (rel.NOME_XLSX, rel.NOME_ZIP) if amb.match(n)])
    esperado_excel = sorted([NOME_XLSX_ESPERADO, "lobo_insights_industrial_v0_1.xlsx", "lobo_insights_industrial_v0_2.xlsx"])
    todos_xlsx = sorted(p.relative_to(ROOT).as_posix() for p in ROOT.rglob("*.xlsx"))
    t("Nomenclatura (V0.2.1)", "excel/ contém só o XLSX atual e os 2 nomes históricos (V0.1.1 e V0.2, não renomeados); nenhum outro .xlsx no projeto", ["excel/" + n for n in esperado_excel], todos_xlsx)
    t("Nomenclatura (V0.2.1)", "os nomes históricos da V0.2 e anteriores foram preservados (XLSX V0.1.1, XLSX V0.2, scripts e documentos das versões anteriores)", [],
      [n for n in list(rel.XLSX_HISTORICOS) if not (EXCEL / n).is_file()] + [n for n in ("scripts/construir_excel_v0_2.py", "scripts/executar_testes_v0_2.py", "docs/V0_2_NOTES.md", "docs/V0_1_1_NOTES.md", "docs/TEST_REPORT_V0_1_1.md")
                                                                              if not (ROOT / n).is_file()])
    t("Nomenclatura (V0.2.1)", "a constante de geração do XLSX (construir_excel_v0_2_1.py) aponta para o nome obrigatório e main() grava nele", "constante correta / main usa NOME_XLSX",
      f"{gen.NOME_XLSX} / {'usa' if 'NOME_XLSX' in inspect.getsource(gen.main) else 'não usa'}", gen.NOME_XLSX == NOME_XLSX_ESPERADO and "NOME_XLSX" in inspect.getsource(gen.main))
    docs = {"README.md": ROOT / "README.md", "CHANGELOG.md": ROOT / "CHANGELOG.md", "docs/V0_2_1_NOTES.md": DOCS / "V0_2_1_NOTES.md"}
    faltam = [f"{k}: {n}" for k, p in docs.items() for n in (NOME_XLSX_ESPERADO, NOME_ZIP_ESPERADO) if n not in p.read_text(encoding="utf-8")]
    t("Nomenclatura (V0.2.1)", "README, CHANGELOG e V0_2_1_NOTES citam os dois nomes obrigatórios exatos", [], faltam)

    # ---------- o XLSX atual é o XLSX auditado da V0.2 (nenhuma alteração analítica/visual)
    h_atual = rel.sha256(XLSX_ATUAL) if XLSX_ATUAL.is_file() else "(XLSX atual ausente)"
    h_v02 = rel.sha256(XLSX_V02) if XLSX_V02.is_file() else "(XLSX histórico da V0.2 ausente)"
    t("Regressão vs V0.2 auditada", "o XLSX atual é byte a byte igual ao XLSX auditado da V0.2 (mesmo SHA-256) ⇒ nenhuma alteração de dados, fórmulas, KPIs, gráficos, layout ou resultados", HASH_XLSX_V02,
      f"{h_atual} / {h_v02}", h_atual == h_v02 == HASH_XLSX_V02 == rel.SHA256_XLSX_V0_2_AUDITADO)
    prot = [l.split("  ", 1) for l in (DOCS / "V0_2_AUDITED_SHA256.txt").read_text(encoding="utf-8").splitlines() if l.strip() and not l.startswith("#")]
    dif = [c for h, c in prot if not (ROOT / c).is_file() or rel.sha256(ROOT / c) != h]
    t("Regressão vs V0.2 auditada", "30 arquivos da V0.2 auditada (CSVs, XLSX históricos, todos os scripts e testes da V0.1.1/V0.2, documentos históricos, notas da V0.2) seguem byte a byte iguais", "30 / 0 diferenças",
      f"{len(prot)} / {len(dif)} diferenças", len(prot) == 30 and not dif)
    obrig = {"data/materiais_ficticios.csv", "data/estoque_ficticio.csv", "data/consumo_ficticio.csv", "scripts/construir_excel_v0_2.py", "scripts/executar_testes_v0_2.py",
             "scripts/construir_excel.py", "scripts/executar_testes.py", "scripts/lobo_common.py", "docs/V0_2_NOTES.md", "excel/lobo_insights_industrial_v0_2.xlsx", "excel/lobo_insights_industrial_v0_1.xlsx"}
    t("Regressão vs V0.2 auditada", "a linha de base da V0.2 inclui dados, regras, XLSX históricos e todo o código analítico e de teste (nada foi excluído da proteção)", [], sorted(obrig - {c for _, c in prot}))
    if not XLSX_ATUAL.is_file():
        t("Regressão vs V0.2 auditada", "o gerador da V0.2.1 reproduz o XLSX atual (fórmulas, valores, gráficos)", "XLSX atual presente", "XLSX atual ausente", False)
    elif localizar_soffice():
        with tempfile.TemporaryDirectory() as tmp:
            novo = Path(tmp) / rel.NOME_XLSX
            info = gen.gerar(novo)
            n, dif_f, dif_v, graf = comparar_xlsx(XLSX_ATUAL, novo)
        t("Regressão vs V0.2 auditada", "o gerador da V0.2.1 (construir_excel_v0_2_1.gerar) reproduz o XLSX atual: mesmas fórmulas/constantes, mesmos valores e mesmos 9 gráficos em todas as células (bytes diferem, conteúdo não)",
          "0 / 0 diferenças / 9 e 9 gráficos / 0 erros", f"{dif_f} / {dif_v} diferenças em {n} células / {graf[0]} e {graf[1]} gráficos / {info['total_errors']} erros",
          dif_f == 0 and dif_v == 0 and graf == (9, 9) and info["total_errors"] == 0)
    else:
        t("Regressão vs V0.2 auditada", "o gerador da V0.2.1 reproduz o XLSX atual (fórmulas, valores, gráficos)", "PASS", "", pular="LibreOffice ausente (restrição do ambiente, não do projeto)")

    # ---------- manifesto operacional SHA256SUMS.txt
    est_txt = (ROOT / rel.MANIFESTO_ESTAVEL).read_text(encoding="utf-8")
    try:
        entradas = rel.ler_manifesto(est_txt)
        formato_ok = True
    except ValueError as e:
        entradas, formato_ok = [], False
        EXTRA["erro_formato"] = str(e)
    caminhos = [c for _, c in entradas]
    t("Manifesto SHA-256 (V0.2.1)", "SHA256SUMS.txt segue o formato do 'sha256sum -c' (sem comentários), está em ordem, sem duplicatas e sem listar a si mesmo, o snapshot nem os regeneráveis", True,
      formato_ok and caminhos == sorted(caminhos, key=lambda s: s.encode("utf-8")) and len(set(caminhos)) == len(caminhos)
      and not ({rel.MANIFESTO_ESTAVEL, rel.MANIFESTO_RELEASE} | set(rel.REGENERAVEIS)) & set(caminhos))
    r = rel.verificar_manifesto(ROOT, rel.MANIFESTO_ESTAVEL)
    t("Manifesto SHA-256 (V0.2.1)", "SHA256SUMS.txt confere integralmente (todos os arquivos estáveis existem e têm o hash registrado)", "0 ausentes / 0 divergentes",
      f"{len(r['ausentes'])} ausentes / {len(r['divergentes'])} divergentes de {r['total']}", not r["ausentes"] and not r["divergentes"])
    if shutil.which("sha256sum"):
        p = subprocess.run(["sha256sum", "-c", rel.MANIFESTO_ESTAVEL], capture_output=True, text=True, cwd=ROOT)
        t("Manifesto SHA-256 (V0.2.1)", "a ferramenta do sistema 'sha256sum -c SHA256SUMS.txt' termina com sucesso", "código 0", f"código {p.returncode}; {p.stdout.count(': OK')} OK; {p.stderr.strip()[:80]}", p.returncode == 0)
    else:
        t("Manifesto SHA-256 (V0.2.1)", "a ferramenta do sistema 'sha256sum -c SHA256SUMS.txt' termina com sucesso", "código 0", "", pular="sha256sum ausente neste sistema (a verificação em Python acima equivale)")
    t("Manifesto SHA-256 (V0.2.1)", "os regeneráveis são exatamente 4 (relatório + 3 prévias), estão fora do manifesto operacional e seguem declarados no módulo de release", REGENERAVEIS_ESPERADOS, list(rel.REGENERAVEIS),
      list(rel.REGENERAVEIS) == REGENERAVEIS_ESPERADOS)
    t("Manifesto SHA-256 (V0.2.1)", "nenhum arquivo do projeto ficou fora do controle: tudo o que não é regenerável nem manifesto está em SHA256SUMS.txt (sem exceções silenciosas)", [], nao_contabilizados(ROOT))
    t("Manifesto SHA-256 (V0.2.1)", "SHA256SUMS.txt é reproduzível: recalculado a partir da árvore atual (gerar_manifesto_estavel) é idêntico ao arquivo gravado", True, rel.gerar_manifesto_estavel(ROOT) == est_txt)
    rel_txt = (ROOT / rel.MANIFESTO_RELEASE).read_text(encoding="utf-8")
    try:
        ent_rel = rel.ler_manifesto(rel_txt)
    except ValueError:
        ent_rel = []
    c_rel = [c for _, c in ent_rel]
    todos = rel.arquivos_do_projeto(ROOT)
    t("Manifesto SHA-256 (V0.2.1)", "RELEASE_SHA256SUMS.txt (snapshot) lista TODOS os arquivos exceto ele mesmo, inclusive SHA256SUMS.txt e os 4 regeneráveis, no formato sha256sum", True,
      bool(ent_rel) and sorted(c_rel) == sorted(c for c in todos if c != rel.MANIFESTO_RELEASE) and rel.MANIFESTO_ESTAVEL in c_rel and set(rel.REGENERAVEIS) <= set(c_rel))
    rr = rel.verificar_manifesto(ROOT, rel.MANIFESTO_RELEASE)
    t("Manifesto SHA-256 (V0.2.1)", "snapshot: nenhum arquivo ausente e nenhuma divergência FORA dos regeneráveis (só os 4 regeneráveis podem diferir do instante do empacotamento; no pacote recém-extraído: 0)",
      "0 ausentes / divergências ⊆ regeneráveis", f"{len(rr['ausentes'])} ausentes / divergentes: {rr['divergentes']}", not rr["ausentes"] and set(rr["divergentes"]) <= set(rel.REGENERAVEIS))

    # ---------- a estratégia distingue evidência regenerável de alteração real (controles positivos e negativos)
    ctrl = []
    with tempfile.TemporaryDirectory() as tmp:
        c = Path(tmp) / "copia"
        copiar_arvore(c)
        for rg in rel.REGENERAVEIS:                                     # execução legítima: os 4 arquivos mudam de bytes
            (c / rg).write_bytes((c / rg).read_bytes() + b"\nvariacao legitima de ambiente\n")
        a = rel.verificar_manifesto(c, rel.MANIFESTO_ESTAVEL)
        b = rel.verificar_manifesto(c, rel.MANIFESTO_RELEASE)
        ctrl.append(("4 regeneráveis alterados", "SHA256SUMS passa", not a["ausentes"] and not a["divergentes"]))
        ctrl.append(("4 regeneráveis alterados", "snapshot acusa exatamente os 4", sorted(b["divergentes"]) == sorted(rel.REGENERAVEIS)))
        (c / "images/dashboard_v0_2_pagina_2.png").unlink()
        a = rel.verificar_manifesto(c, rel.MANIFESTO_ESTAVEL)
        ctrl.append(("1 imagem regenerável apagada", "SHA256SUMS passa", not a["ausentes"] and not a["divergentes"]))
        for alvo in ("data/consumo_ficticio.csv", "scripts/construir_excel_v0_2.py", "excel/" + NOME_XLSX_ESPERADO, "docs/V0_2_NOTES.md"):   # alterações reais
            if not (c / alvo).is_file():
                ctrl.append((f"alteração real em {alvo}", "arquivo-alvo do controle ausente", False))
                continue
            orig = (c / alvo).read_bytes()
            (c / alvo).write_bytes(orig + b" ")
            a = rel.verificar_manifesto(c, rel.MANIFESTO_ESTAVEL)
            ctrl.append((f"alteração real em {alvo}", "SHA256SUMS reprova só esse arquivo", a["divergentes"] == [alvo] and not a["ausentes"]))
            (c / alvo).write_bytes(orig)
        (c / "scripts/lobo_common.py").unlink()
        a = rel.verificar_manifesto(c, rel.MANIFESTO_ESTAVEL)
        ctrl.append(("arquivo estável apagado", "SHA256SUMS acusa ausente", a["ausentes"] == ["scripts/lobo_common.py"]))
        (c / "scripts/lobo_common.py").write_bytes((ROOT / "scripts/lobo_common.py").read_bytes())
        (c / "scripts/novo_nao_registrado.py").write_text("x = 1\n", encoding="utf-8")
        ctrl.append(("arquivo novo não registrado", "teste de cobertura o detecta", nao_contabilizados(c) == ["scripts/novo_nao_registrado.py"]))
    EXTRA["controles"] = ctrl
    t("Manifesto SHA-256 (V0.2.1)", "controle (cópia descartável): execução legítima que muda os 4 regeneráveis — ou apaga um deles — NÃO invalida SHA256SUMS.txt, mas o snapshot registra a mudança",
      "3 / 3 verificações", f"{sum(k for _, _, k in ctrl[:3])} / 3 verificações", all(k for _, _, k in ctrl[:3]))
    t("Manifesto SHA-256 (V0.2.1)", "controle (cópia descartável): alteração REAL num CSV, num script, no XLSX ou numa nota, arquivo estável apagado e arquivo novo não registrado são detectados (nada é mascarado)",
      "6 / 6 detecções", f"{sum(k for _, _, k in ctrl[3:])} / 6 detecções", all(k for _, _, k in ctrl[3:]) and len(ctrl[3:]) == 6)

    # ---------- documentação da V0.2.1
    ch = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    t("Documentação V0.2.1", "CHANGELOG registra 'V0.2.1 — Packaging & Reproducibility Fix' acima de V0.2, V0.1.1 e V0.1 (históricos preservados)", True,
      "V0.2.1 — Packaging & Reproducibility Fix" in ch and ch.index("V0.2.1 — Packaging & Reproducibility Fix") < ch.index("V0.2 — Analytics & Dashboard") < ch.index("V0.1.1 — Portable Baseline") < ch.index("V0.1 — Foundation"))
    nt = (DOCS / "V0_2_1_NOTES.md").read_text(encoding="utf-8")
    sec = ["Resultado da auditoria da V0.2", "Ausência de alteração funcional", "Correção de nomenclatura", "Estratégia de hashes", "Arquivos regeneráveis", "Testes executados", "Confirmação de regressão", "Confirmação dos KPIs"]
    t("Documentação V0.2.1", "V0_2_1_NOTES.md contém o aviso de dados sintéticos, as 8 seções exigidas, os dois manifestos e os 4 arquivos regeneráveis",
      [], [s for s in sec if s not in nt] + [x for x in (rel.MANIFESTO_ESTAVEL, rel.MANIFESTO_RELEASE, *rel.REGENERAVEIS) if x not in nt] + ([] if AVISO in nt else ["aviso"]))
    rd = (ROOT / "README.md").read_text(encoding="utf-8")
    t("Documentação V0.2.1", "README cita V0.2.1, os dois manifestos, o executor da V0.2.1 e continua com a tabela dos 6 KPIs e o histórico (XLSX V0.1/V0.2)", [],
      [x for x in ("V0.2.1", rel.MANIFESTO_ESTAVEL, rel.MANIFESTO_RELEASE, "executar_testes_v0_2_1.py", "lobo_insights_industrial_v0_2.xlsx", "lobo_insights_industrial_v0_1.xlsx") if x not in rd]
      + [n for n, v in base.kpis_csv().items() if not re.search(rf"\|\s*{re.escape(n)}\s*\|", rd)])
    t("Documentação V0.2.1", "docs/manual/observacoes_execucao_v0_2_1.md existe e registra o que foi (e não foi) reproduzido do problema dos hashes", True,
      (DOCS / "manual" / "observacoes_execucao_v0_2_1.md").is_file() and "regener" in (DOCS / "manual" / "observacoes_execucao_v0_2_1.md").read_text(encoding="utf-8").lower())


# ================================================================== relatório
def relatorio(tab_q, resumo_legado, n_legado, n_v02):
    def md(cab, rows):
        out = ["| " + " | ".join(cab) + " |", "|" + "|".join("---" for _ in cab) + "|"]
        out += ["| " + " | ".join(base.esc(c) for c in r) + " |" for r in rows]
        return "\n".join(out)
    exe = localizar_soffice()
    try:
        lo = subprocess.run([exe, "--version"], capture_output=True, text=True, timeout=60).stdout.strip() if exe else "LibreOffice NÃO encontrado"
    except Exception:
        lo = "LibreOffice (versão não obtida)"
    L = base.LINHAS
    cont = lambda sel, s: sum(1 for x in sel if x[5] == s)      # noqa: E731
    novos = L[n_v02:]
    ex = v02.EXTRA
    ins = (DOCS / "manual" / "inspecao_visual_v0_2.md").read_text(encoding="utf-8")
    obs = (DOCS / "manual" / "observacoes_execucao_v0_2.md").read_text(encoding="utf-8")
    obs1 = (DOCS / "manual" / "observacoes_execucao_v0_2_1.md").read_text(encoding="utf-8")
    ctrl = EXTRA["controles"]
    return "\n".join([
        "# TEST_REPORT — Projeto Lobo 1 — Lobo Insights Industrial V0.2.1 (Packaging & Reproducibility Fix)",
        f"> {AVISO}",
        "",
        "Relatório **gerado por `scripts/executar_testes_v0_2_1.py`**; cada PASS/FAIL vem de uma execução real e teste que não pôde rodar aparece como **SKIP** (nunca PASS). "
        "**Este arquivo é regenerável**: registra o ambiente de execução (versões e plataforma) e por isso não faz parte do manifesto operacional `SHA256SUMS.txt` (ver `docs/V0_2_1_NOTES.md`). "
        "A V0.2.1 é uma microversão de empacotamento: **nenhuma lógica analítica, dado, fórmula, gráfico ou layout mudou** — o XLSX atual é byte a byte igual ao XLSX auditado da V0.2. "
        "Históricos preservados: `docs/TEST_REPORT_V0_1.md`, `docs/TEST_REPORT_V0_1_1.md` (o relatório da V0.2 auditada está descrito em `docs/V0_2_NOTES.md`). Para reproduzir: `python scripts/executar_testes_v0_2_1.py --verificar`.",
        "",
        "## 1. Resumo",
        f"- Testes automatizados: **{len(L)}** — PASS **{cont(L, 'PASS')}**, FAIL **{cont(L, 'FAIL')}**, SKIP **{cont(L, 'SKIP')}**.",
        f"- Suíte da V0.1.1 (T01–T{n_legado}, sem alteração): {resumo_legado}",
        f"- Bateria da V0.2 (T{n_legado + 1}–T{n_v02}, sem alteração, **executada sobre `excel/{NOME_XLSX_ESPERADO}`**): **{n_v02 - n_legado}** — PASS **{cont(L[n_legado:n_v02], 'PASS')}**, FAIL **{cont(L[n_legado:n_v02], 'FAIL')}**, SKIP **{cont(L[n_legado:n_v02], 'SKIP')}**.",
        f"- Testes novos da V0.2.1 (T{n_v02 + 1}–T{len(L)}): **{len(novos)}** — PASS **{cont(novos, 'PASS')}**, FAIL **{cont(novos, 'FAIL')}**, SKIP **{cont(novos, 'SKIP')}**.",
        f"- Verificações de qualidade dos CSVs (Python, Q01–Q36): **{len(tab_q)}** — PASS **{int((tab_q['Status'] == 'PASS').sum())}**, FAIL **{int((tab_q['Status'] == 'FAIL').sum())}**.",
        f"- Ambiente: Python {platform.python_version()}, pandas {pd.__version__}, openpyxl {openpyxl.__version__}, {lo}, {platform.system()} {platform.machine()}.",
        f"- SHA-256 do XLSX atual (`excel/{NOME_XLSX_ESPERADO}`): `{rel.sha256(XLSX_ATUAL)}`",
        f"- SHA-256 do XLSX auditado da V0.2 (`excel/lobo_insights_industrial_v0_2.xlsx`, nome histórico preservado): `{rel.sha256(XLSX_V02)}` — **idêntico**.",
        "- SHA-256 dos CSVs (inalterados): " + "; ".join(f"`{n}` = `{rel.sha256(DATA / n)}`" for n in base.ARQ),
        "",
        "## 2. TESTADO × NÃO TESTADO",
        "**TESTADO (executado):** toda a suíte da V0.1.1; toda a bateria da V0.2 (KPIs, 36 verificações de qualidade, 2.417 fórmulas sem erro, sem vínculos externos, 9 gráficos, cálculo independente de todas as tabelas da aba Analises, mutação, reconstrução, renderização) **agora sobre o XLSX com o nome da V0.2.1**; "
        "nomenclatura obrigatória (XLSX, ZIP, rejeição de nomes ambíguos, nomes históricos preservados); igualdade byte a byte do XLSX atual com o auditado; 30 arquivos da V0.2 auditada inalterados; "
        "manifesto operacional (formato, ordem, cobertura, reprodutibilidade, `sha256sum -c`); snapshot da release; **controles positivos e negativos da estratégia de hashes** em cópia descartável; documentação.",
        "",
        "**NÃO TESTADO (declarado, não assumido):**",
        "- **Microsoft Excel real** (Windows/Mac/Online) e **Excel em pt-BR**: mesmas limitações da V0.2 (`docs/LIMITATIONS.md`, itens 17–25). Como o XLSX é byte a byte o mesmo da V0.2, nada novo foi verificado nem quebrado neste ponto.",
        "- **Variação real do ambiente do auditor:** a falha de manifesto relatada **não foi reproduzida** aqui (nesta máquina as regenerações saem byte a byte idênticas); a estratégia foi validada por simulação — os 4 regeneráveis foram alterados numa cópia (ver seção 3) — e pela execução completa dos testes seguida de `sha256sum -c`, registradas em `docs/manual/observacoes_execucao_v0_2_1.md`.",
        "- **O ZIP final** só pode ser conferido após o empacotamento (um arquivo não contém a prova do próprio ZIP): essa conferência (nome, conteúdo, `RELEASE_SHA256SUMS.txt`, execução a partir de outro diretório) é feita fora desta suíte e registrada nas observações.",
        "- Instalação limpa por `pip install -r requirements.txt` (sem rede neste ambiente); Windows/macOS; outras versões de Python/LibreOffice.",
        "- Inspeção visual humana: nenhuma alteração visual foi feita; vale a inspeção da V0.2 (seção 8), parcial.",
        "",
        "## 3. Manifestos SHA-256 e arquivos regeneráveis",
        "- `SHA256SUMS.txt` — manifesto **operacional**: todos os arquivos estáveis (dados, scripts, XLSX, documentação-fonte, manifestos históricos), sem os 4 regeneráveis. Continua válido depois de executar os testes.",
        "- `RELEASE_SHA256SUMS.txt` — **snapshot** de todos os arquivos (inclusive `SHA256SUMS.txt` e os regeneráveis) no instante do empacotamento; só confere integralmente no pacote recém-extraído.",
        "- Regeneráveis: " + ", ".join(f"`{x}`" for x in rel.REGENERAVEIS) + ".",
        "",
        "Controles da estratégia (cópia descartável do projeto):",
        "",
        md(["Situação simulada", "Verificação", "Resultado"], [(a, b, "PASS" if k else "FAIL") for a, b, k in ctrl]),
        "",
        "## 4. Verificações de qualidade dos dados (Python, sobre os CSVs)",
        md(["Teste", "Regra", "Resultado", "Quantidade de erros", "Status"], [(f"{r.ID} · {r.Teste}", r.Regra, r.Resultado, r["Quantidade de erros"], r.Status) for _, r in tab_q.iterrows()]),
        "",
        "## 5. Testes automatizados (T01–T103 V0.1.1; T104–T174 V0.2; T175 em diante V0.2.1)",
        md(["ID", "Área", "Cenário", "Esperado", "Obtido", "PASS/FAIL"], L),
        "",
        "## 6. Regressão dos 6 KPIs oficiais",
        md(["KPI", "Referência", "XLSX V0.1.1", "Indicadores (XLSX atual)", "Cartão do DASHBOARD", "CSV independente", "Status"],
           [(n, e, a, b, c, d, "PASS" if ok else "FAIL") for n, e, a, b, c, d, ok in ex["regressao"]]),
        "",
        "## 7. Inventário dos gráficos e teste de mutação (executados sobre o XLSX atual)",
        md(["Arquivo", "Tipo", "Direção", "Séries", "Pontos", "Rótulos", "Referências"],
           [(g["arquivo"].split("/")[-1], g["tipos"], g["dir"], g["series"], g["pontos"], g["rotulos"], "; ".join(g["refs"])) for g in ex["inventario_graficos"]]),
        "",
        (md(["Indicador (mutação)", "Dados oficiais", "Esperado com dados alterados", "Obtido no Excel recalculado"], ex["mutacao"]) if ex["mutacao"] else "Teste de mutação **NÃO EXECUTADO** (LibreOffice ausente)."),
        "",
        "## 8. Inspeção visual (da V0.2; sem alteração visual na V0.2.1)",
        ins,
        "",
        "## 9. Observações de execução",
        "### V0.2",
        obs,
        "",
        "### V0.2.1",
        obs1,
        "",
    ])


def main(escrever):
    v02.XLSX02 = XLSX_ATUAL                          # a bateria da V0.2 passa a rodar sobre o XLSX com o nome da V0.2.1
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        codigo_legado = base.main(escrever=False)
    saida = buf.getvalue().strip().splitlines()
    n_legado = len(base.LINHAS)
    v02.EXTRA["n_legado"] = n_legado
    resumo = next((s for s in saida if s.startswith("Testes automatizados")), "")
    resumo += " | " + " | ".join(s for s in saida if s.startswith(("Verificações", "Detecção")))
    print("SUÍTE LEGADA V0.1.1:", resumo, f"(código de saída {codigo_legado})")
    if XLSX_ATUAL.is_file():
        dfs, _, _ = v02.testes_v0_2()
    else:                                            # sem o XLSX atual a bateria da V0.2 não tem o que testar: registra a falha em vez de quebrar
        dfs = validar_dados.carregar()
        t("Nomenclatura (V0.2.1)", "o XLSX atual existe com o nome obrigatório (a bateria da V0.2 não pôde ser executada sem ele)", NOME_XLSX_ESPERADO, "arquivo ausente", False)
    n_v02 = len(base.LINHAS)
    print(f"Bateria V0.2 sobre o XLSX atual: {n_v02 - n_legado} testes | PASS {sum(x[5] == 'PASS' for x in base.LINHAS[n_legado:])} | FAIL {sum(x[5] == 'FAIL' for x in base.LINHAS[n_legado:])} | SKIP {sum(x[5] == 'SKIP' for x in base.LINHAS[n_legado:])}")
    testes_v0_2_1()
    L = base.LINHAS
    novos = L[n_v02:]
    falhas, pulados = [x for x in L if x[5] == "FAIL"], [x for x in L if x[5] == "SKIP"]
    print(f"Testes novos da V0.2.1: {len(novos)} | PASS {sum(x[5] == 'PASS' for x in novos)} | FAIL {sum(x[5] == 'FAIL' for x in novos)} | SKIP {sum(x[5] == 'SKIP' for x in novos)}")
    print(f"TOTAL: {len(L)} | PASS {len(L) - len(falhas) - len(pulados)} | FAIL {len(falhas)} | SKIP {len(pulados)}")
    for x in falhas:
        print("FAIL:", x)
    for x in pulados:
        print("SKIP:", x[0], x[2][:90], "->", x[4])
    if escrever and XLSX_ATUAL.is_file():
        (DOCS / "TEST_REPORT.md").write_text(relatorio(validar_dados.tabela(validar_dados.executar(dfs)), resumo, n_legado, n_v02), encoding="utf-8")
        print("TEST_REPORT.md escrito; prévias:", v02.gerar_previas())
    if falhas or codigo_legado == 1:
        return 1
    return 2 if (pulados or codigo_legado == 2) else 0


if __name__ == "__main__":
    sys.exit(main(escrever="--verificar" not in sys.argv))
