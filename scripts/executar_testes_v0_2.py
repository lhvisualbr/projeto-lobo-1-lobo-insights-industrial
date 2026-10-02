# -*- coding: utf-8 -*-
"""Suíte de testes da V0.2 (Analytics & Dashboard). Tudo aqui é EXECUTADO de verdade.

Uso:
    python executar_testes_v0_2.py              # roda tudo e (re)escreve docs/TEST_REPORT.md e images/dashboard_v0_2_*.png
    python executar_testes_v0_2.py --verificar  # roda tudo, NÃO escreve relatório nem imagens

Estratégia: a suíte da V0.1.1 (`executar_testes.py`, T01-T103) é importada e executada SEM NENHUMA alteração;
os testes novos (T104 em diante) continuam a numeração. Teste que não puder rodar (ex.: sem LibreOffice) vira SKIP,
nunca PASS. Códigos de saída: 0 = tudo executado e PASS; 1 = há FAIL; 2 = nenhum FAIL, mas há SKIP.

Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.
"""
import contextlib
import hashlib
import io
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import openpyxl
import pandas as pd
from lxml import etree

import executar_testes as base
import validar_dados
from construir_excel import NOME_XLSX, localizar_soffice, recalcular
from construir_excel_v0_2 import NOME_XLSX_V02, S_ANA, S_DASH, construir_v0_2
from lobo_common import CATEGORIAS, CENTROS, DATA, DOCS, EXCEL, ROOT, status_estoque

t = base.t
AVISO = base.AVISO
XLSX02 = EXCEL / NOME_XLSX_V02
XLSX01 = EXCEL / NOME_XLSX
REF = {"Unidades consumidas": 1647, "Custo estimado consumido": Decimal("24606.20"), "Movimentações": 180,
       "Materiais distintos movimentados": 20, "Itens abaixo ou iguais ao mínimo": 5,
       "Itens críticos abaixo ou iguais ao mínimo": 2}          # referências de regressão fixadas no enunciado da V0.2
NS_C = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
EXTRA = {"inventario_graficos": [], "mutacao": [], "render": {}, "regressao": [], "sem_lo": None}


def sha256(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


# ================================================================== cálculo independente (Python puro + Decimal)
def _rows(df):
    return df.to_dict("records")


def ref_calc(dfs):
    """Recalcula TUDO o que a aba Analises mostra, só a partir dos dataframes (sem ler nenhuma fórmula do Excel)."""
    mat, est, con = _rows(dfs["mat"]), _rows(dfs["est"]), _rows(dfs["con"])
    cons = [r for r in con if r["Tipo_Movimentacao"] == "Consumo"]
    q = lambda r: int(r["Quantidade"])                                              # noqa: E731
    c = lambda r: Decimal(r["Quantidade"]) * Decimal(r["Custo_Unitario"])          # noqa: E731
    tot_un, tot_cu = sum(q(r) for r in cons), sum(c(r) for r in cons)

    def grupo(chaves, campo):
        out = []
        for k in chaves:
            sel = [r for r in cons if r[campo] == k]
            out.append({"k": k, "mov": len(sel), "un": sum(q(r) for r in sel), "cu": sum(c(r) for r in sel)})
        return out

    def ordenar(itens, campo):
        idx = sorted(range(len(itens)), key=lambda i: (-itens[i][campo], i))       # desempate: ordem original
        return [itens[i] for i in idx]

    meses = ["2026-06", "2026-07", "2026-08"]
    R = {"tot_un": tot_un, "tot_cu": tot_cu, "mov": len(cons), "mat_dist": len({r["Material_ID"] for r in cons})}
    R["mes"] = [
        {"k": m, "mov": sum(1 for r in cons if r["Data"][:7] == m), "un": sum(q(r) for r in cons if r["Data"][:7] == m),
         "cu": sum(c(r) for r in cons if r["Data"][:7] == m)} for m in meses]
    R["cat"], R["cen"] = grupo(CATEGORIAS, "Categoria"), grupo(CENTROS, "Centro_Trabalho")
    R["cat_un"], R["cat_cu"] = ordenar(R["cat"], "un"), ordenar(R["cat"], "cu")
    R["cen_un"], R["cen_cu"] = ordenar(R["cen"], "un"), ordenar(R["cen"], "cu")
    mats = []
    for m in mat:
        sel = [r for r in cons if r["Material_ID"] == m["Material_ID"]]
        mats.append({"k": m["Material_ID"], "nome": m["Material"], "un": sum(q(r) for r in sel), "cu": sum(c(r) for r in sel)})
    for campo, tot in (("cu", tot_cu), ("un", tot_un)):
        ordem, acum, lista = ordenar(mats, campo), Decimal(0), []
        for m in ordem:
            pct = (Decimal(m[campo]) / Decimal(tot)) if tot else Decimal(0)
            antes = acum
            acum += pct
            lista.append({"k": m["k"], "nome": m["nome"], "v": m[campo], "pct": pct, "acum": acum, "nucleo": antes < Decimal("0.8")})
        R["par_" + campo] = lista
        pos = [x for x in lista if x["v"] > 0]
        R["n80_" + campo] = min(len(pos), sum(1 for x in lista if x["acum"] < Decimal("0.8")) + 1)
        R["top3_" + campo], R["top5_" + campo] = lista[2]["acum"], lista[4]["acum"]
    # estoque
    fila = []
    for i, e in enumerate(est):
        a, mi = int(e["Estoque_Atual"]), int(e["Estoque_Minimo"])
        s = status_estoque(a, mi)
        cr = {"Alta": 1, "Média": 2, "Baixa": 3}[e["Criticidade"]]
        razao = Decimal(a) / Decimal(mi) if mi else Decimal("9.999")
        chave = {"REPOR": 1, "ATENÇÃO": 2, "OK": 3}[s] * 1000000 + cr * 10000 + min(9999, int((razao * 1000).to_integral_value(ROUND_HALF_UP)))
        fila.append({"i": i, "id": e["Material_ID"], "nome": e["Material"], "crit": e["Criticidade"], "atual": a, "min": mi,
                     "status": s, "risco": e["Criticidade"] == "Alta" and s == "REPOR", "chave": chave})
    R["fila"] = sorted(fila, key=lambda x: (x["chave"], x["i"]))
    R["matriz"] = {cr: {s: sum(1 for x in fila if x["crit"] == cr and x["status"] == s) for s in ("REPOR", "ATENÇÃO", "OK")}
                   for cr in ("Alta", "Média", "Baixa")}
    R["repor"] = sum(x["status"] == "REPOR" for x in fila)
    R["risco"] = sum(x["risco"] for x in fila)
    return R


def near(a, b, tol=0.005):
    return abs(float(a) - float(b)) <= tol


# ================================================================== leitura das camadas do XLSX V0.2
def posicoes(dfs):
    """Posições das camadas novas: o layout não depende dos dados, então basta construir o arquivo cru (sem recalcular)."""
    with tempfile.TemporaryDirectory() as tmp:
        return construir_v0_2(dfs, Path(tmp) / "pos.xlsx", {})


def conferir_analises(vals, P, R):
    """Compara TODAS as tabelas da aba Analises (valores) com o cálculo independente. Retorna lista de divergências."""
    a = vals[S_ANA]
    g = lambda ref: a.get(ref)                                                      # noqa: E731
    div = []

    def cmp(rotulo, x, y, tol=0.005):
        ok = near(x, y, tol) if isinstance(y, (int, float, Decimal)) and not isinstance(y, bool) and isinstance(x, (int, float)) else x == y
        if not ok:
            div.append(f"{rotulo}: Excel={x!r} esperado={y!r}")

    m = P["mes"]
    for k, x in enumerate(R["mes"]):
        r = m["r0"] + k
        cmp(f"mês {x['k']} mov", g(f"C{r}"), x["mov"]); cmp(f"mês {x['k']} un", g(f"D{r}"), x["un"]); cmp(f"mês {x['k']} custo", g(f"E{r}"), x["cu"])
        if k:
            ant = R["mes"][k - 1]
            cmp(f"mês {x['k']} var un", g(f"F{r}"), float(Decimal(x["un"]) / Decimal(ant["un"]) - 1), 1e-9)
            cmp(f"mês {x['k']} var custo", g(f"G{r}"), float(x["cu"] / ant["cu"] - 1), 1e-9)
    for nome, bl, ordem_u, ordem_c in (("cat", P["cat"], R["cat_un"], R["cat_cu"]), ("cen", P["cen"], R["cen_un"], R["cen_cu"])):
        for k in range(len(ordem_u)):
            r = bl["s0"] + k
            cmp(f"{nome} un pos{k + 1} nome", g(f"B{r}"), ordem_u[k]["k"]); cmp(f"{nome} un pos{k + 1} valor", g(f"C{r}"), ordem_u[k]["un"])
            cmp(f"{nome} un pos{k + 1} %", g(f"D{r}"), float(Decimal(ordem_u[k]["un"]) / Decimal(R["tot_un"])), 1e-9)
            cmp(f"{nome} custo pos{k + 1} nome", g(f"G{r}"), ordem_c[k]["k"]); cmp(f"{nome} custo pos{k + 1} valor", g(f"H{r}"), ordem_c[k]["cu"])
            cmp(f"{nome} custo pos{k + 1} %", g(f"I{r}"), float(ordem_c[k]["cu"] / R["tot_cu"]), 1e-9)
    for chave, bl in (("par_cu", P["par_c"]), ("par_un", P["par_u"])):
        for k, x in enumerate(R[chave]):
            r = bl["p0"] + k
            cmp(f"{chave} pos{k + 1} id", g(f"B{r}"), x["k"]); cmp(f"{chave} pos{k + 1} material", g(f"C{r}"), x["nome"])
            cmp(f"{chave} pos{k + 1} valor", g(f"D{r}"), x["v"]); cmp(f"{chave} pos{k + 1} %", g(f"E{r}"), float(x["pct"]), 1e-9)
            cmp(f"{chave} pos{k + 1} acum", g(f"F{r}"), float(x["acum"]), 1e-9)
            cmp(f"{chave} pos{k + 1} faixa", g(f"H{r}"), "Núcleo" if x["nucleo"] else "Cauda")
            cmp(f"{chave} pos{k + 1} núcleo", g(f"I{r}"), x["v"] if x["nucleo"] else 0)
            cmp(f"{chave} pos{k + 1} demais", g(f"J{r}"), 0 if x["nucleo"] else x["v"])
    co = P["conc"]
    cmp("materiais p/ 80% custo", g(f"B{co['mat80_c']}"), R["n80_cu"]); cmp("materiais p/ 80% unidades", g(f"B{co['mat80_u']}"), R["n80_un"])
    cmp("top3 custo", g(f"B{co['topa_c']}"), float(R["top3_cu"]), 1e-9); cmp("top5 custo", g(f"B{co['topb_c']}"), float(R["top5_cu"]), 1e-9)
    cmp("top3 unid", g(f"B{co['topa_u']}"), float(R["top3_un"]), 1e-9); cmp("top5 unid", g(f"B{co['topb_u']}"), float(R["top5_un"]), 1e-9)
    cmp("maior categoria custo", g(f"C{co['cat_c']}"), R["cat_cu"][0]["k"]); cmp("maior centro unid", g(f"C{co['cen_u']}"), R["cen_un"][0]["k"])
    q = P["est_q"]
    for k, x in enumerate(R["fila"]):
        r = q["q0"] + k
        cmp(f"fila pos{k + 1} id", g(f"B{r}"), x["id"]); cmp(f"fila pos{k + 1} status", g(f"J{r}"), x["status"])
        cmp(f"fila pos{k + 1} atual", g(f"F{r}"), x["atual"]); cmp(f"fila pos{k + 1} mínimo", g(f"G{r}"), x["min"])
        cmp(f"fila pos{k + 1} risco", g(f"K{r}"), "Sim" if x["risco"] else "Não")
    pn = P["pnl"]
    fora = [x for x in R["fila"] if x["status"] != "OK"]
    for k in range(10):
        r = pn["n0"] + k
        if k < len(fora):
            cmp(f"painel pos{k + 1} material", g(f"B{r}"), fora[k]["nome"]); cmp(f"painel pos{k + 1} status", g(f"F{r}"), fora[k]["status"])
        else:
            cmp(f"painel pos{k + 1} vazio", g(f"B{r}") or "", "")
    cmp("itens fora do OK", g(f"B{pn['fora']}"), len(fora)); cmp("itens não exibidos", g(f"B{pn['oculto']}"), max(0, len(fora) - 10))
    mt = P["mtx"]
    for k, cr in enumerate(("Alta", "Média", "Baixa")):
        for col, s in (("B", "REPOR"), ("C", "ATENÇÃO"), ("D", "OK")):
            cmp(f"matriz {cr}×{s}", g(f"{col}{mt['x0'] + k}"), R["matriz"][cr][s])
    cf = P["conf"]
    ok = sum(g(f"E{r}") == "OK" for r in range(cf["c0"], cf["c1"] + 1))
    cmp("conferências OK", ok, cf["c1"] - cf["c0"] + 1)
    return div


def conferir_dashboard(vals, P, R, dash):
    """Cartões de KPI, mini-cartões, fila de atenção e matriz do DASHBOARD contra o cálculo independente."""
    d, div = vals[S_DASH], []
    esperados = [R["tot_un"], R["tot_cu"], R["mov"], R["mat_dist"], R["repor"], R["risco"]]
    for col, e in zip("BDFHJL", esperados):
        v = d.get(f"{col}{dash['kpi_row']}")
        if not (isinstance(v, (int, float)) and near(v, e)):
            div.append(f"cartão {col}: {v!r} esperado {e!r}")
    l0, _ = dash["estoque_lista"]
    fora = [x for x in R["fila"] if x["status"] != "OK"]
    for k in range(10):
        nome, stt = d.get(f"B{l0 + k}") or "", d.get(f"G{l0 + k}") or ""
        en, es = (fora[k]["nome"], fora[k]["status"]) if k < len(fora) else ("", "")
        if (nome, stt) != (en, es):
            div.append(f"fila dashboard linha {k + 1}: ({nome!r},{stt!r}) esperado ({en!r},{es!r})")
    return div


# ================================================================== inspeção dos gráficos (XML)
def inventario_graficos(xlsx, P, vals):
    """Lê o XML de cada gráfico: tipo, séries, referências (c:f) e pontos; confere se cada referência aponta para células preenchidas."""
    out, problemas = [], []
    a = vals[S_ANA]
    with zipfile.ZipFile(xlsx) as z:
        nomes = sorted((n for n in z.namelist() if re.fullmatch(r"xl/charts/chart\d+\.xml", n)), key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
        for n in nomes:
            raiz = etree.fromstring(z.read(n))
            tipos = [e.tag.split("}")[1] for e in raiz.find("c:chart/c:plotArea", NS_C) if e.tag.endswith("Chart")]
            direc = raiz.xpath("string(//c:barChart/c:barDir/@val)", namespaces=NS_C)
            n_ser = len(raiz.findall(".//c:ser", NS_C))
            refs = [e.text for e in raiz.iterfind(".//c:f", NS_C)]
            titulo_auto = raiz.find("c:chart/c:autoTitleDeleted", NS_C) is not None and raiz.find("c:chart/c:title", NS_C) is None
            npts = None
            for ref in refs:
                m = re.fullmatch(r"'?(\w+)'?!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)", ref)
                if not m:
                    m = re.fullmatch(r"'?(\w+)'?!([A-Z]+)(\d+)()()", ref)          # título da série: célula única, sem $
                if not m or m.group(1) != S_ANA:
                    problemas.append(f"{n}: referência inesperada {ref}")
                    continue
                col1, r1 = m.group(2), int(m.group(3))
                r2 = int(m.group(5)) if m.group(5) else r1
                cel = [a.get(f"{col1}{r}") for r in range(r1, r2 + 1)]
                if r1 != r2:
                    npts = r2 - r1 + 1
                if any(c is None for c in cel) and not n.endswith("chart9.xml"):
                    problemas.append(f"{n}: {ref} contém célula vazia")
            with_val = [e.text for e in raiz.iterfind(".//c:val//c:f", NS_C)]
            for ref in with_val:
                m = re.fullmatch(r"'?\w+'?!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)", ref or "")
                if not m:
                    problemas.append(f"{n}: referência de valores inválida ou quebrada: {ref}")
                    continue
                cel = [a.get(f"{m.group(1)}{r}") for r in range(int(m.group(2)), int(m.group(4)) + 1)]
                if not all(isinstance(c, (int, float)) or c in ("", None) for c in cel):
                    problemas.append(f"{n}: valores não numéricos em {ref}")
            out.append({"arquivo": n, "tipos": "+".join(tipos), "dir": direc or "-", "series": n_ser, "pontos": npts,
                        "sem_titulo_auto": titulo_auto, "refs": refs, "eixos_val": len(raiz.findall(".//c:valAx", NS_C)),
                        "rotulos": len(raiz.findall(".//c:dLbls", NS_C))})
    return out, problemas


def esperado_graficos(P):
    """(tipo, direção, nº de séries, nº de pontos, intervalos de valores, intervalo de categorias, células de título das séries)."""
    m, c, e, pc, pu, pn = P["mes"], P["cat"], P["cen"], P["par_c"], P["par_u"], P["pnl"]
    rng = lambda col, r1, r2: f"'{S_ANA}'!${col}${r1}:${col}${r2}"                  # noqa: E731
    cel = lambda col, r: f"'{S_ANA}'!{col}{r}"                                      # noqa: E731
    return [
        ("barChart", "col", 1, 3, [rng("D", m["r0"], m["r1"])], rng("B", m["r0"], m["r1"]), []),
        ("barChart", "col", 1, 3, [rng("E", m["r0"], m["r1"])], rng("B", m["r0"], m["r1"]), []),
        ("barChart", "bar", 1, 6, [rng("C", c["s0"], c["s1"])], rng("B", c["s0"], c["s1"]), []),
        ("barChart", "bar", 1, 6, [rng("H", c["s0"], c["s1"])], rng("G", c["s0"], c["s1"]), []),
        ("barChart", "bar", 1, 5, [rng("C", e["s0"], e["s1"])], rng("B", e["s0"], e["s1"]), []),
        ("barChart", "bar", 1, 5, [rng("H", e["s0"], e["s1"])], rng("G", e["s0"], e["s1"]), []),
        ("barChart+lineChart", "col", 4, 20, [rng(x, pc["p0"], pc["p1"]) for x in "IJFG"], rng("C", pc["p0"], pc["p1"]), [cel(x, pc["hdr"]) for x in "IJFG"]),
        ("barChart+lineChart", "col", 4, 20, [rng(x, pu["p0"], pu["p1"]) for x in "IJFG"], rng("C", pu["p0"], pu["p1"]), [cel(x, pu["hdr"]) for x in "IJFG"]),
        ("barChart", "bar", 2, 10, [rng("D", pn["n0"], pn["n1"]), rng("E", pn["n0"], pn["n1"])], rng("B", pn["n0"], pn["n1"]), [cel("D", pn["n0"] - 1), cel("E", pn["n0"] - 1)]),
    ]


# ================================================================== PDF / imagem (inspeção automatizada)
def renderizar(xlsx, pasta):
    """Converte o XLSX em PDF (LibreOffice) e as 3 primeiras páginas em PPM (pdftoppm). Retorna (pdf, [ppm])."""
    soffice = localizar_soffice()
    perfil = Path(pasta) / "perfil"
    (perfil / "user").mkdir(parents=True)
    entrada = Path(pasta) / "entrada.xlsx"
    shutil.copy(xlsx, entrada)
    subprocess.run([soffice, "--headless", "--norestore", f"-env:UserInstallation={perfil.as_uri()}", "--convert-to", "pdf",
                    "--outdir", str(pasta), str(entrada)], capture_output=True, text=True, timeout=240)
    pdf = Path(pasta) / "entrada.pdf"
    if not pdf.exists():
        raise RuntimeError("LibreOffice não gerou o PDF")
    subprocess.run(["pdftoppm", "-r", "50", "-f", "1", "-l", "3", str(pdf), str(Path(pasta) / "pg")], capture_output=True, timeout=120)
    return pdf, sorted(Path(pasta).glob("pg-*.ppm")) or sorted(Path(pasta).glob("pg-*.png"))


def contar_cor(ppm, alvo, tol=28):
    """Conta pixels próximos de uma cor (RGB) num arquivo PPM (P6), em Python puro."""
    dados = Path(ppm).read_bytes()
    partes = dados.split(b"\n", 3)
    larg, alt = (int(x) for x in partes[1].split())
    pix = partes[3]
    n = 0
    for i in range(0, larg * alt * 3, 3):
        if abs(pix[i] - alvo[0]) < tol and abs(pix[i + 1] - alvo[1]) < tol and abs(pix[i + 2] - alvo[2]) < tol:
            n += 1
    return n


# ================================================================== mutação: prova de que nada no dashboard é fixo
def dados_mutados(dfs):
    mat, est, con = (dfs[k].copy() for k in ("mat", "est", "con"))
    con.loc[con["Material_ID"] == "MAT010", "Tipo_Movimentacao"] = "Ajuste"        # MAT010 deixa de ter consumo (KPI 4 = 19)
    idx = con.index[(con["Material_ID"] == "MAT019")][0]
    con.loc[idx, "Quantidade"] = "300"                                             # MAT019 (custo 19,00) sobe no ranking de unidades
    idx2 = con.index[(con["Material_ID"] == "MAT004")][0]
    con.loc[idx2, "Data"] = "2026-08-15"                                           # muda a distribuição mensal
    for mid, atual in (("MAT008", "50"), ("MAT016", "41"), ("MAT001", "40"), ("MAT005", "40")):
        est.loc[est["Material_ID"] == mid, "Estoque_Atual"] = atual                # MAT005 sai de REPOR; MAT008 e MAT001 entram; MAT016 vira ATENÇÃO
    return {"mat": mat, "est": est, "con": con}


# ================================================================== testes
def testes_v0_2():
    dfs = validar_dados.carregar()
    ref = ref_calc(dfs)
    wf = openpyxl.load_workbook(XLSX02)
    wv = openpyxl.load_workbook(XLSX02, data_only=True)
    w1f, w1v = openpyxl.load_workbook(XLSX01), openpyxl.load_workbook(XLSX01, data_only=True)
    vals = {ws.title: {c.coordinate: c.value for row in ws.iter_rows() for c in row if c.value is not None} for ws in wv.worksheets}
    P_all = posicoes(dfs)
    P, dash = P_all["P"], P_all["dash"]

    # ---------- regressão contra a baseline V0.1.1 (arquivos protegidos)
    base_txt = (DOCS / "V0_1_1_BASELINE_SHA256.txt").read_text(encoding="utf-8")
    prot = [l.split("  ", 1) for l in base_txt.splitlines() if l.strip() and not l.startswith("#")]
    hist = {"docs/TEST_REPORT.md": "docs/TEST_REPORT_V0_1_1.md"}                   # o relatório da V0.1.1 foi preservado com outro nome
    diverg = [rel for h, rel in prot if not (ROOT / hist.get(rel, rel)).exists() or sha256(ROOT / hist.get(rel, rel)) != h]
    t("Regressão vs V0.1.1", "arquivos protegidos da V0.1.1 (CSVs, XLSX V0.1, scripts e regras, testes, DATA_DICTIONARY, notas e relatórios históricos) seguem byte a byte iguais", [], diverg)
    t("Regressão vs V0.1.1", "a base V0.1.1 registra 23 arquivos protegidos, incluindo os 3 CSVs, o XLSX histórico e a suíte executar_testes.py", 23, len(prot),
      len(prot) == 23 and {"data/consumo_ficticio.csv", "excel/lobo_insights_industrial_v0_1.xlsx", "scripts/executar_testes.py", "scripts/construir_excel.py"} <= {r for _, r in prot})
    t("Regressão vs V0.1.1", "os 3 CSVs oficiais têm os SHA-256 da baseline informados no enunciado da V0.2 (sem alteração de dados)", [],
      [n for n, h in (("materiais_ficticios.csv", "eca98597dc740956b4db95885ec54693def431a56501ea6d055fa54c1654231b"),
                      ("estoque_ficticio.csv", "9f90cd52f208f9f863b2ababf7fed4735bbe893ef137efdc79f7506b3359a8d0"),
                      ("consumo_ficticio.csv", "b5b8249731e5fcbce34b4345b4d1e576c568b33de3eba4dfca91f8bd0d94233c")) if sha256(DATA / n) != h])
    t("Regressão vs V0.1.1", "XLSX histórico V0.1.1 (lobo_insights_industrial_v0_1.xlsx) preservado e distinto do XLSX V0.2", True,
      XLSX01.exists() and XLSX02.exists() and sha256(XLSX01) != sha256(XLSX02) and XLSX01.name != XLSX02.name)

    # ---------- KPIs oficiais: referências fixas x CSV independente x Indicadores x DASHBOARD
    ind = wv["Indicadores"]
    kp_csv = base.kpis_csv()
    for (nome, esperado), col in zip(REF.items(), "BDFHJL"):
        r = base.achar(ind, nome)
        v_ind, v_dash = ind.cell(r, 2).value, vals[S_DASH].get(f"{col}{dash['kpi_row']}")
        v_01 = w1v["Indicadores"].cell(base.achar(w1v["Indicadores"], nome), 2).value
        ok = near(v_ind, esperado) and near(v_dash, esperado) and near(v_01, esperado) and near(kp_csv[nome], esperado)
        EXTRA["regressao"].append((nome, esperado, v_01, v_ind, v_dash, kp_csv[nome], ok))
        t("KPIs oficiais (regressão)", f"'{nome}': referência {esperado} = XLSX V0.1.1 = Indicadores V0.2 = cartão do DASHBOARD = CSV independente", str(esperado),
          f"{v_01} / {v_ind} / {v_dash} / {kp_csv[nome]}", ok)
    t("KPIs oficiais (regressão)", "rótulos dos 6 cartões do DASHBOARD = nomes oficiais dos KPIs (Indicadores)", list(REF), [vals[S_DASH].get(f"{c}6") for c in "BDFHJL"])

    # ---------- estrutura do XLSX V0.2
    t("Excel V0.2 – estrutura", "abas na ordem (DASHBOARD primeiro)", ["DASHBOARD", "LEIA_ME", "Indicadores", "Analises", "Materiais", "Estoque", "Consumo", "Qualidade", "Parametros"], wf.sheetnames)
    t("Excel V0.2 – estrutura", "DASHBOARD é a aba ativa ao abrir (única selecionada)", "DASHBOARD / [DASHBOARD]", f"{wf.active.title} / {[ws.title for ws in wf.worksheets if ws.sheet_view.tabSelected]}",
      wf.active.title == "DASHBOARD" and [ws.title for ws in wf.worksheets if ws.sheet_view.tabSelected] == ["DASHBOARD"])
    tabs = {n: ws.tables[n].ref for ws in wf.worksheets for n in ws.tables}
    t("Excel V0.2 – estrutura", "tabelas estruturadas e intervalos idênticos aos da V0.1.1", {"tbMateriais": "A1:G21", "tbEstoque": "A1:L21", "tbConsumo": "A1:L181", "tbQualidade": "A9:E45"}, tabs)
    t("Excel V0.2 – estrutura", "nomes definidos = 9 da V0.1.1 + pLimitePareto, pTopA, pTopB",
      sorted(["lstCategorias", "lstUnidades", "lstCriticidade", "lstCentros", "lstTipoMov", "lstMeses", "pMargemAtencao", "pInicioPeriodo", "pFimPeriodo", "pLimitePareto", "pTopA", "pTopB"]), sorted(wf.defined_names.keys()))
    with zipfile.ZipFile(XLSX02) as z:
        partes, wbxml = z.namelist(), z.read("xl/workbook.xml").decode("utf-8")
        xml_todo = "\n".join(z.read(n).decode("utf-8", "ignore") for n in partes if n.endswith((".xml", ".rels", ".vml")))
    n_ext, n_rel_ext = sum(p.startswith("xl/externalLinks") for p in partes), xml_todo.count("TargetMode=" + chr(34) + "External" + chr(34))
    t("Excel V0.2 – estrutura", "sem vínculos externos (externalLinks) e sem relações externas (TargetMode=External)", "0 / 0", f"{n_ext} / {n_rel_ext}", n_ext == 0 and n_rel_ext == 0)
    t("Excel V0.2 – estrutura", "recálculo total ao abrir (fullCalcOnLoad)", True, "fullCalcOnLoad" in wbxml)
    amb = ["/m" + "nt/", "/ho" + "me/", "/t" + "mp/", "C:" + "\\", "file:" + "//", "Users" + "\\"]
    t("Excel V0.2 – estrutura", "nenhum caminho absoluto/privado dentro do XLSX (fórmulas, gráficos, hiperlinks, propriedades)", [], [a for a in amb if a in xml_todo])
    hl = [(c.coordinate, c.hyperlink.location, c.hyperlink.target) for row in wf[S_DASH].iter_rows() for c in row if c.hyperlink]
    t("Excel V0.2 – estrutura", "6 hiperlinks do DASHBOARD são internos (location) e nenhum aponta para fora do arquivo", "6 / 0 externos",
      f"{len(hl)} / {sum(1 for _, loc, tg in hl if tg or not loc)} externos", len(hl) == 6 and all(loc and not tg for _, loc, tg in hl))
    destinos = [loc.split("!")[0] for _, loc, _ in hl]
    t("Excel V0.2 – estrutura", "destinos dos hiperlinks existem (abas do próprio arquivo)", [], [d for d in destinos if d not in wf.sheetnames])

    # ---------- preservação das abas de dados / indicadores / parâmetros
    def dif_abas(nome, permitidas=()):
        a, b, d = w1f[nome], wf[nome], []
        coords = {c.coordinate for row in a.iter_rows() for c in row if c.value is not None} | {c.coordinate for row in b.iter_rows() for c in row if c.value is not None}
        for co in coords:
            if co in permitidas:
                continue
            if a[co].value != b[co].value:
                d.append(co)
        return d

    for nome in ("Materiais", "Estoque", "Consumo", "Qualidade"):
        dv_a, dv_b = len(w1f[nome].data_validations.dataValidation), len(wf[nome].data_validations.dataValidation)
        t("Preservação da base", f"aba {nome}: todas as células (constantes e fórmulas) idênticas à V0.1.1; validações de dados e formatação condicional iguais", "0 diferenças",
          f"{len(dif_abas(nome))} diferenças; DV {dv_a}/{dv_b}; CF {len(w1f[nome].conditional_formatting)}/{len(wf[nome].conditional_formatting)}",
          not dif_abas(nome) and dv_a == dv_b and len(w1f[nome].conditional_formatting) == len(wf[nome].conditional_formatting))
    t("Preservação da base", "aba Indicadores: fórmulas idênticas à V0.1.1; únicas diferenças = título A1 e nota A3 (texto)", ["A1", "A3"], sorted(dif_abas("Indicadores", ())))
    todas_dif = dif_abas("Parametros")
    dif_par = [c for c in todas_dif if int(re.sub(r"[A-Z]+", "", c)) <= 18]
    novos_par = [wf["Parametros"].cell(r, 2).value for r in (22, 23, 24)]
    permitido = {col + str(lin) for col in "ABC" for lin in range(20, 25)}
    fora_bloco = [c for c in todas_dif if c not in permitido]
    t("Preservação da base", "aba Parametros: A1:F18 idênticas à V0.1.1; todas as diferenças estão em A20:C24 (Limite_Pareto=0,8; Top_N_A=3; Top_N_B=5)",
      "0 diferenças fora de A20:C24 / [0.8, 3, 5]", f"{len(fora_bloco) + len(dif_par)} diferenças / {novos_par}", not dif_par and not fora_bloco and novos_par == [0.8, 3, 5])
    vals01 = {ws.title: {c.coordinate: c.value for row in ws.iter_rows() for c in row if c.value is not None} for ws in w1v.worksheets}
    dif_v = []
    for nome in ("Materiais", "Estoque", "Consumo", "Qualidade", "Indicadores"):
        for co, v in vals01[nome].items():
            if co in ("A1", "A3") and nome == "Indicadores":
                continue
            w = vals[nome].get(co)
            if not ((near(v, w, 1e-9) if isinstance(v, (int, float)) and isinstance(w, (int, float)) else v == w)):
                dif_v.append(f"{nome}!{co}")
    t("Preservação da base", "valores calculados de Materiais/Estoque/Consumo/Qualidade/Indicadores = V0.1.1 (todas as células)", 0, len(dif_v))

    # ---------- fórmulas
    formulas, sem_cache, erros, const_num = 0, [], [], {n: [] for n in (S_DASH, S_ANA, "Indicadores")}
    inventario = {}
    for ws in wf.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    formulas += 1
                    for fn in re.findall(r"([A-Z][A-Z0-9.]*)\(", c.value):
                        inventario[fn] = inventario.get(fn, 0) + 1
                    v = wv[ws.title][c.coordinate].value
                    if v is None and ws.title not in ("DASHBOARD", "Analises"):
                        sem_cache.append(f"{ws.title}!{c.coordinate}")
                    elif isinstance(v, str) and v.startswith("#"):
                        erros.append(f"{ws.title}!{c.coordinate}")
                elif ws.title in const_num and isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                    const_num[ws.title].append(c.coordinate)
    t("Excel V0.2 – fórmulas", "fórmulas reais no arquivo (> 2000; V0.1.1 tinha 857)", "> 2000", formulas, formulas > 2000)
    t("Excel V0.2 – fórmulas", "células com erro (#REF!, #DIV/0!, #VALUE!, #NAME?, #N/A…) nos valores calculados: todas as abas", 0, len(erros))
    xml_erros = re.findall(r"#(?:REF|DIV/0|VALUE|NAME|N/A|NUM|NULL)[!?]?", xml_todo)
    t("Excel V0.2 – fórmulas", "nenhum '#REF!' ou intervalo quebrado dentro do XML (fórmulas, nomes definidos, gráficos)", 0, len([x for x in xml_erros if x.startswith("#REF")]))
    sc = []
    with zipfile.ZipFile(XLSX02) as z:
        for n in z.namelist():
            if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n):
                raiz = etree.fromstring(z.read(n))
                for cel in raiz.iterfind(".//{*}c"):
                    if cel.find("{*}f") is not None and cel.find("{*}v") is None:
                        sc.append(f"{n}:{cel.get('r')}")
    t("Excel V0.2 – fórmulas", "toda célula de fórmula tem valor em cache gravado (<v>) — visualizadores e celular mostram os números", 0, len(sc))
    t("Excel V0.2 – fórmulas", "nenhum número digitado (constante numérica) nas abas DASHBOARD, Analises e Indicadores — tudo é fórmula", "0 / 0 / 0",
      f"{len(const_num[S_DASH])} / {len(const_num[S_ANA])} / {len(const_num['Indicadores'])}", not any(const_num.values()))
    prox = ["FORE" + "CAST", "TR" + "END", "GRO" + "WTH", "LIN" + "EST", "LOG" + "EST", "SLO" + "PE", "INTER" + "CEPT"]
    todas = " ".join(c.value for ws in wf.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("=")).upper()
    t("Escopo", "nenhuma função de previsão/tendência/regressão (" + ", ".join(prox) + ") nas fórmulas do XLSX V0.2", [], [w for w in prox if re.search(rf"\b{w}(\.\w+)?\(", todas)])
    nao_port = ["XLOOKUP", "XMATCH", "SORT", "FILTER", "UNIQUE", "SEQUENCE", "LET", "LAMBDA"]
    t("Compatibilidade", "nenhuma função de matriz dinâmica/pós-2019 (XLOOKUP, SORT, FILTER, UNIQUE, SEQUENCE, LET, LAMBDA) — ordenações por RANK+COUNTIF+INDEX/MATCH", [], [w for w in nao_port if re.search(rf"(?<![A-Z.]){w}\(", todas)])
    refs_dash = set()
    for row in wf[S_DASH].iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                refs_dash |= set(re.findall(r"([A-Za-z_]\w*)!", c.value))
    t("Auditabilidade", "fórmulas do DASHBOARD só leem as abas Analises, Indicadores e Qualidade (cadeia dados → Indicadores → Analises → DASHBOARD)", ["Analises", "Indicadores", "Qualidade"], sorted(refs_dash))
    refs_ana = set()
    for row in wf[S_ANA].iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                refs_ana |= set(re.findall(r"([A-Za-z_]\w*)!", c.value))
    t("Auditabilidade", "fórmulas da aba Analises só leem Indicadores, Parametros, tabelas estruturadas e o DASHBOARD (apenas na conferência dos cartões)", ["DASHBOARD", "Indicadores", "Parametros"], sorted(refs_ana))

    # ---------- camada analítica x cálculo independente
    div = conferir_analises(vals, P, ref)
    t("Análises – valores", "TODAS as tabelas da aba Analises (mensal, categoria, centro, Pareto de custo e de consumo, concentração, fila de estoque, painel, matriz, conferências) = cálculo independente dos CSVs (Decimal)", "0 divergências", f"{len(div)} divergências", not div)
    if div:
        EXTRA["divergencias"] = div[:10]
    t("Análises – valores", "resultado esperado dos dados oficiais: 9 materiais somam 80% do custo; 4 materiais somam 80% das unidades", "9 / 4", f"{ref['n80_cu']} / {ref['n80_un']}", (ref["n80_cu"], ref["n80_un"]) == (9, 4))
    ties = [x["v"] for x in ref["par_un"]]
    t("Análises – valores", "há empates reais nas unidades por material (MAT011/MAT012/MAT020 = 11; MAT006/MAT018 = 5) e o desempate é a ordem do catálogo", "MAT011,MAT012,MAT020 / MAT006,MAT018",
      ",".join(x["k"] for x in ref["par_un"] if x["v"] == 11) + " / " + ",".join(x["k"] for x in ref["par_un"] if x["v"] == 5),
      len(ties) != len(set(ties)) and [x["k"] for x in ref["par_un"] if x["v"] == 11] == ["MAT011", "MAT012", "MAT020"])
    t("Análises – valores", "Pareto: % acumulado é não decrescente e termina em 100%; soma dos valores = KPI (custo e unidades)", True,
      all(ref[k][i]["acum"] <= ref[k][i + 1]["acum"] for k in ("par_cu", "par_un") for i in range(19)) and ref["par_cu"][-1]["acum"] == 1 and ref["par_un"][-1]["acum"] == 1)
    resumo_conf = vals[S_ANA].get("B" + str(P["conf"]["resumo"]))
    t("Análises – valores", "conferências internas da aba Analises (seção 8) e de Indicadores (seção 6)", "18 de 18 / 12 de 12", f"{resumo_conf} / {ind.cell(16, 2).value}",
      resumo_conf == "18 de 18" and ind.cell(16, 2).value == "12 de 12")
    div_d = conferir_dashboard(vals, P, ref, dash)
    t("DASHBOARD – valores", "cartões de KPI e fila de atenção do DASHBOARD = cálculo independente dos CSVs", "0 divergências", f"{len(div_d)} divergências", not div_d)
    txt = " | ".join(str(v) for v in vals[S_DASH].values() if isinstance(v, str))
    esperados = ["Jul/2026", "691 un.", "Corte", "940 un.", "9 de 20", "4 de 20", "MECANICA", "MANUTENCAO", "5 em REPOR (2 de criticidade Alta)", "4 em ATENÇÃO (2 de criticidade Alta)",
                 "18 de 18 OK", "36 PASS / 0 FAIL", "12 de 12 OK", "Disco de Corte 4,5 pol é o maior em custo"]
    t("DASHBOARD – valores", "textos de leitura gerados por fórmula trazem os fatos esperados dos dados (mês de maior consumo, maiores categorias/centros, concentração, contagens de estoque, conferências)", [], [e for e in esperados if e not in txt])
    t("DASHBOARD – valores", "variações mensais nas leituras: +13,7% (jul/jun) e -49,6% (ago/jul) para unidades; +29,3% e -52,3% para custo (separador decimal do LibreOffice ou do Excel)", [],
      [x for x in (r"\+13[.,]7%", r"-49[.,]6%", r"\+29[.,]3%", r"-52[.,]3%") if not re.search(x, txt)])
    lim = ["Dados 100% sintéticos", "sem previsão", "não é recomendação operacional", "não realiza previsão"]
    t("Escopo", "DASHBOARD declara: dados sintéticos, sem previsão e sem recomendação operacional", [], [x for x in lim if x not in txt])
    prev = re.findall(r"(?i)\b(previs[ãa]o de|proje[çc][ãa]o|tend[êe]ncia (?:futura|de alta|de queda)|deve[rá]+ (?:subir|cair))\b", txt.replace("sem previsão", "").replace("não realiza previsão", "").replace("não indicam tendência", ""))
    t("Escopo", "DASHBOARD não contém linguagem de previsão/projeção (fora das negações explícitas)", [], prev)

    # ---------- gráficos
    inv, problemas = inventario_graficos(XLSX02, P, vals)
    EXTRA["inventario_graficos"] = inv
    esp = esperado_graficos(P)
    t("Gráficos", "9 gráficos no DASHBOARD (6 de barras simples, 2 Pareto combinados, 1 comparativo de estoque) e nenhum em outra aba", "9 / 0", f"{len(inv)} / {sum(len(ws._charts) for ws in wf.worksheets if ws.title != S_DASH)}",
      len(inv) == 9 and sum(len(ws._charts) for ws in wf.worksheets if ws.title != S_DASH) == 0)
    falhas = []
    for k, (g, e) in enumerate(zip(inv, esp), start=1):
        tipo, direc, nser, npts, vals_ref, cat_ref, tit_ref = e
        obtido_val = [r for r in g["refs"] if r in vals_ref]
        obtido_tit = [r for r in g["refs"] if r in tit_ref]
        if not (g["tipos"] == tipo and g["dir"] == direc and g["series"] == nser and g["pontos"] == npts and obtido_val == vals_ref and cat_ref in g["refs"] and obtido_tit == tit_ref):
            falhas.append(f"gráfico {k}: {g['tipos']}/{g['dir']}/{g['series']}/{g['pontos']} refs={g['refs']}")
    t("Gráficos", "cada gráfico: tipo, orientação, nº de séries, nº de pontos e intervalos (valores e categorias) = o esperado da camada Analises", [], falhas)
    t("Gráficos", "todas as referências apontam só para a aba Analises, em células preenchidas e numéricas (nenhum #REF!, nenhuma célula vazia)", [], problemas)
    t("Gráficos", "todos os gráficos declaram 'sem título automático' e têm rótulos de dados; Paretos têm eixo secundário (2 eixos de valor)", True,
      all(g["sem_titulo_auto"] for g in inv) and all(g["rotulos"] >= 1 for g in inv if g["tipos"] == "barChart") and all(g["eixos_val"] == 2 for g in inv if "lineChart" in g["tipos"]))
    with zipfile.ZipFile(XLSX02) as z:
        ancoras = etree.fromstring(z.read("xl/drawings/drawing1.xml")).findall("{*}twoCellAnchor")
    t("Gráficos", "os 9 gráficos estão ancorados por células (twoCellAnchor) dentro da área impressa do DASHBOARD", 9, len(ancoras), len(ancoras) == 9)

    # ---------- layout, impressão, painéis
    d = wf[S_DASH]
    t("Layout / impressão", "DASHBOARD: paisagem, A4, largura ajustada a 1 página, área de impressão definida, 2 quebras manuais de página, sem linhas de grade",
      "landscape/9/1/2/sem grade", f"{d.page_setup.orientation}/{d.page_setup.paperSize}/{d.page_setup.fitToWidth}/{len(d.row_breaks.brk)}/{'com' if d.sheet_view.showGridLines else 'sem'} grade",
      d.page_setup.orientation == "landscape" and d.page_setup.paperSize == 9 and d.page_setup.fitToWidth == 1 and d.print_area and len(d.row_breaks.brk) == 2 and not d.sheet_view.showGridLines)
    larg = [d.column_dimensions[c].width for c in "BCDEFGHIJKLM"]
    t("Layout / impressão", "DASHBOARD: 12 colunas de mesma largura (grade uniforme) e sem congelamento de painéis (decisão: rolagem livre em notebook)", "12 x 13 / sem freeze", f"{set(larg)} / {d.freeze_panes}", set(larg) == {13} and d.freeze_panes is None)
    t("Layout / impressão", "abas de dados congelam o cabeçalho (Materiais/Estoque/Consumo A2; Qualidade A10; Indicadores A4; Analises A6)", "A2/A2/A2/A10/A4/A6",
      "/".join(str(wf[n].freeze_panes) for n in ("Materiais", "Estoque", "Consumo", "Qualidade", "Indicadores", "Analises")),
      [wf[n].freeze_panes for n in ("Materiais", "Estoque", "Consumo", "Qualidade", "Indicadores", "Analises")] == ["A2", "A2", "A2", "A10", "A4", "A6"])
    t("Layout / impressão", "Analises: paisagem e largura ajustada; Analises sem linhas de grade", True, wf[S_ANA].page_setup.orientation == "landscape" and wf[S_ANA].page_setup.fitToWidth == 1 and not wf[S_ANA].sheet_view.showGridLines)
    cf_dash = len(d.conditional_formatting)
    t("Layout / impressão", "DASHBOARD: formatação condicional nos status de estoque, criticidade, matriz e cartões de alerta (>= 6 intervalos)", ">= 6", cf_dash, cf_dash >= 6)
    fontes = {c.font.name for row in d.iter_rows() for c in row if c.value is not None}
    t("Layout / impressão", "DASHBOARD usa uma única família tipográfica (Arial)", {"Arial"}, fontes)
    alturas = [d.row_dimensions[r].height for r in (6, 7, 8)]
    t("Layout / impressão", "cartões de KPI: linhas de rótulo, valor e leitura com altura definida (27/38/28 pt) para o texto não ser cortado", [27, 38, 28], alturas)
    t("Layout / impressão", "cartões de KPI: formatos numéricos (unidades #,##0; custo #,##0.00; contagens #,##0)", ["#,##0", "#,##0.00", "#,##0", "#,##0", "#,##0", "#,##0"], [d[f"{c}7"].number_format for c in "BDFHJL"])

    # ---------- reconstrução idêntica + inspeção visual + mutação (dependem do LibreOffice)
    lo = localizar_soffice()
    motivo = "LibreOffice ausente (restrição do ambiente, não do projeto)"
    if not lo:
        EXTRA["sem_lo"] = motivo
        for cen in ("reconstruir o XLSX V0.2 a partir dos CSVs reproduz o XLSX entregue (fórmulas e valores)",
                    "teste de mutação: dados alterados => Analises e DASHBOARD recalculam e batem com o cálculo independente",
                    "teste de mutação: fila de estoque com mais itens que o painel (overflow) e ranking de unidades alterado"):
            t("Reconstrução / mutação", cen, "PASS", "", pular=motivo)
    else:
        with tempfile.TemporaryDirectory() as tmp:
            cru = Path(tmp) / "cru.xlsx"
            construir_v0_2(dfs, cru, {"sha": {p.name: sha256(p) for p in sorted(DATA.glob("*.csv"))}})
            novo_val, novo_info = recalcular(cru)
            wn = openpyxl.load_workbook(cru)
            n_cel = dif_f = dif_v2 = 0
            for ws in wf.worksheets:
                for row in ws.iter_rows():
                    for c in row:
                        if c.value is None:
                            continue
                        n_cel += 1
                        dif_f += wn[ws.title][c.coordinate].value != c.value
                        if isinstance(c.value, str) and c.value.startswith("="):
                            x, y = wv[ws.title][c.coordinate].value, novo_val[ws.title].get(c.coordinate)
                            x = None if x == "" else x
                            dif_v2 += not ((abs(x - y) < 1e-9) if isinstance(x, (int, float)) and isinstance(y, (int, float)) and not isinstance(x, bool) else x == y)
            wn.close()
        t("Reconstrução / mutação", "reconstruir o XLSX V0.2 a partir dos CSVs (construir_excel_v0_2.py + recálculo) reproduz o XLSX entregue: fórmulas/constantes e valores de todas as células", "0 / 0 diferenças",
          f"{dif_f} / {dif_v2} diferenças em {n_cel} células ({novo_info['total_formulas']} fórmulas, {novo_info['total_errors']} erros)", dif_f == 0 and dif_v2 == 0 and novo_info["total_errors"] == 0)
        # mutação
        dfm = dados_mutados(dfs)
        refm = ref_calc(dfm)
        with tempfile.TemporaryDirectory() as tmp:
            arq = Path(tmp) / "mutado.xlsx"
            construir_v0_2(dfm, arq, {})
            vm, im = recalcular(arq)
        divm = conferir_analises(vm, P, refm)
        divm_d = conferir_dashboard(vm, P, refm, dash)
        EXTRA["mutacao"] = [
            ("KPI Materiais distintos movimentados", ref["mat_dist"], refm["mat_dist"], vm[S_DASH].get(f"H{dash['kpi_row']}")),
            ("KPI Itens abaixo ou iguais ao mínimo (REPOR)", ref["repor"], refm["repor"], vm[S_DASH].get(f"J{dash['kpi_row']}")),
            ("KPI Itens críticos (Alta em REPOR)", ref["risco"], refm["risco"], vm[S_DASH].get(f"L{dash['kpi_row']}")),
            ("KPI Unidades consumidas", ref["tot_un"], refm["tot_un"], vm[S_DASH].get(f"B{dash['kpi_row']}")),
            ("Materiais que somam 80% das unidades", ref["n80_un"], refm["n80_un"], vm[S_ANA].get(f"B{P['conc']['mat80_u']}")),
            ("2º do ranking de unidades (material)", ref["par_un"][1]["k"], refm["par_un"][1]["k"], vm[S_ANA].get("B" + str(P["par_u"]["p0"] + 1))),
            ("Itens fora do OK (estoque)", len([x for x in ref["fila"] if x["status"] != "OK"]), len([x for x in refm["fila"] if x["status"] != "OK"]), vm[S_ANA].get(f"B{P['pnl']['fora']}")),
            ("Itens não exibidos no painel", max(0, len([x for x in ref["fila"] if x["status"] != "OK"]) - 10), max(0, len([x for x in refm["fila"] if x["status"] != "OK"]) - 10), vm[S_ANA].get(f"B{P['pnl']['oculto']}")),
        ]
        t("Reconstrução / mutação", "teste de mutação: dados alterados (consumo de MAT010 vira 'Ajuste', MAT019 = 300 un., MAT004 muda de mês, estoques de MAT001/005/008/016 alterados) => TODAS as tabelas da Analises e os cartões/fila do DASHBOARD recalculam e batem com o cálculo independente dos dados alterados",
          "0 / 0 divergências (0 erros)", f"{len(divm)} / {len(divm_d)} divergências ({im['total_errors']} erros)", not divm and not divm_d and im["total_errors"] == 0)
        mudou = [x for x in EXTRA["mutacao"] if x[1] != x[2]]
        t("Reconstrução / mutação", "a mutação é efetiva: os resultados esperados MUDAM em relação aos dados oficiais (KPI 4 = 19, ranking, REPOR, overflow do painel)", ">= 6 indicadores diferentes", f"{len(mudou)} de {len(EXTRA['mutacao'])} diferentes",
          len(mudou) >= 6 and all(x[2] == x[3] or (isinstance(x[3], (int, float)) and near(x[3], x[2])) for x in EXTRA["mutacao"]))
        # inspeção visual automatizada
        if shutil.which("pdftoppm") and shutil.which("pdfinfo"):
            with tempfile.TemporaryDirectory() as tmp:
                pdf, pgs = renderizar(XLSX02, tmp)
                txt_p = [subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), "-layout", str(pdf), "-"], capture_output=True, text=True).stdout for i in (1, 2, 3)]
                pixels = {}
                for i, pg in enumerate(pgs[:3], start=1):
                    if pg.suffix == ".ppm":
                        pixels[i] = (contar_cor(pg, (0x2F, 0x66, 0x90)), contar_cor(pg, (0xC9, 0x77, 0x2B)))
                n_pag = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True).stdout).group(1))
            EXTRA["render"] = {"paginas_total": n_pag, "pixels": pixels}
            t("Inspeção visual (automatizada)", "PDF do workbook gerado; as 3 primeiras páginas são o DASHBOARD (cabeçalho e KPIs; Pareto; estoque)", "LOBO INSIGHTS / Pareto de custo / Fila de atenção",
              " / ".join(x for x, k in (("LOBO INSIGHTS", "LOBO INSIGHTS INDUSTRIAL" in txt_p[0]), ("Pareto de custo", "Pareto de custo por material" in txt_p[1]), ("Fila de atenção", "Fila de atenção" in txt_p[2])) if k) or "nenhum",
              "LOBO INSIGHTS INDUSTRIAL" in txt_p[0] and "Pareto de custo por material" in txt_p[1] and "Fila de atenção" in txt_p[2])
            if pixels:
                t("Inspeção visual (automatizada)", "gráficos realmente desenhados: pixels das cores de série (azul de unidades e cobre de custo) presentes nas páginas 1 e 2 e barras de estoque na 3", "azul>2000 e cobre>2000 (p1 e p2); azul>500 (p3)",
                  f"p1 {pixels[1]}; p2 {pixels[2]}; p3 {pixels[3]}", pixels[1][0] > 2000 and pixels[1][1] > 2000 and pixels[2][0] > 2000 and pixels[2][1] > 2000 and pixels[3][0] > 500)
            else:
                t("Inspeção visual (automatizada)", "gráficos desenhados (contagem de pixels)", "PASS", "", pular="pdftoppm não gerou PPM")
        else:
            for cen in ("PDF do workbook gerado; as 3 primeiras páginas são o DASHBOARD", "gráficos realmente desenhados (contagem de pixels)"):
                t("Inspeção visual (automatizada)", cen, "PASS", "", pular="pdftoppm/pdfinfo (poppler) ausentes")

    # ---------- documentação da V0.2
    ok_arq = ["docs/V0_2_NOTES.md", "docs/V0_1_1_BASELINE_SHA256.txt", "docs/TEST_REPORT_V0_1_1.md", "docs/manual/inspecao_visual_v0_2.md", "docs/manual/observacoes_execucao_v0_2.md",
              "scripts/construir_excel_v0_2.py", "scripts/executar_testes_v0_2.py", f"excel/{NOME_XLSX_V02}"]
    t("Documentação V0.2", "arquivos novos da V0.2 presentes", [], [a for a in ok_arq if not (ROOT / a).exists()])
    nt = (DOCS / "V0_2_NOTES.md").read_text(encoding="utf-8")
    sec = ["Objetivo", "Recursos adicionados", "Decisões analíticas", "Decisões visuais", "Limitações", "Testes executados", "Evidências de regressão"]
    t("Documentação V0.2", "V0_2_NOTES.md contém aviso de dados sintéticos e as 7 seções exigidas", [], [s for s in sec if s not in nt] + ([] if AVISO in nt else ["aviso"]))
    ch = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    t("Documentação V0.2", "CHANGELOG registra 'V0.2 — Analytics & Dashboard' acima de V0.1.1 e V0.1 (históricos preservados)", True,
      "V0.2 — Analytics & Dashboard" in ch and ch.index("V0.2 — Analytics & Dashboard") < ch.index("V0.1.1 — Portable Baseline") < ch.index("V0.1 — Foundation"))
    rd = (ROOT / "README.md").read_text(encoding="utf-8")
    kp = base.kpis_csv()
    faltam = [f"{n}={base.fmt_money(v) if isinstance(v, Decimal) else base.fmt_int(v)}" for n, v in kp.items()
              if not re.search(rf"\|\s*{re.escape(n)}\s*\|\s*{re.escape(base.fmt_money(v) if isinstance(v, Decimal) else base.fmt_int(v))}\s*\|", rd)]
    t("Documentação V0.2", "README cita 'V0.2', o XLSX V0.2, o DASHBOARD e os 6 KPIs com os valores reais", [], faltam + [x for x in ("V0.2", NOME_XLSX_V02, "DASHBOARD") if x not in rd])
    lim = (DOCS / "LIMITATIONS.md").read_text(encoding="utf-8").lower()
    t("Documentação V0.2", "LIMITATIONS declara: sintético, sem previsão, sem recomendação operacional e as limitações novas da V0.2 (Excel real, filtros, cache/pré-visualização)", [],
      [x for x in ("sintétic", "previsão", "uso operacional", "microsoft excel", "filtros", "v0.2") if x not in lim])
    rel = base.LINHAS
    t("Documentação V0.2", "docs/TEST_REPORT_V0_1_1.md é o relatório histórico da V0.1.1 (título e 103 testes)", True,
      (DOCS / "TEST_REPORT_V0_1_1.md").read_text(encoding="utf-8").startswith("# TEST_REPORT — Lobo Insights Industrial V0.1.1") and "103" in (DOCS / "TEST_REPORT_V0_1_1.md").read_text(encoding="utf-8")[:1500])
    t("Documentação V0.2", "a suíte legada foi executada sem alteração: executar_testes.py idêntico ao da V0.1.1 e 103 testes legados registrados antes dos novos", "idêntico / 103", f"{'idêntico' if sha256(ROOT / 'scripts' / 'executar_testes.py') == 'f4bfd733cf271a99f38e734a07786805c3ddf0a4fa1a402b080a77434bfa71e7' else 'ALTERADO'} / {EXTRA.get('n_legado')}",
      sha256(ROOT / "scripts" / "executar_testes.py") == "f4bfd733cf271a99f38e734a07786805c3ddf0a4fa1a402b080a77434bfa71e7" and EXTRA.get("n_legado") == 103)
    return dfs, ref, rel


# ================================================================== relatório
def esc(s):
    return base.esc(s)


def relatorio(tab_q, wv, dfs, resumo_legado, n_legado):
    def md(cab, rows):
        out = ["| " + " | ".join(cab) + " |", "|" + "|".join("---" for _ in cab) + "|"]
        out += ["| " + " | ".join(esc(c) for c in r) + " |" for r in rows]
        return "\n".join(out)
    exe = localizar_soffice()
    try:
        lo = subprocess.run([exe, "--version"], capture_output=True, text=True, timeout=60).stdout.strip() if exe else "LibreOffice NÃO encontrado"
    except Exception:
        lo = "LibreOffice (versão não obtida)"
    L = base.LINHAS
    npass, nfail, nskip = (sum(l[5] == s for l in L) for s in ("PASS", "FAIL", "SKIP"))
    ins = (DOCS / "manual" / "inspecao_visual_v0_2.md").read_text(encoding="utf-8")
    obs = (DOCS / "manual" / "observacoes_execucao_v0_2.md").read_text(encoding="utf-8")
    q_wb = [(wv["Qualidade"].cell(10 + k, 1).value, wv["Qualidade"].cell(10 + k, 4).value, wv["Qualidade"].cell(10 + k, 5).value) for k in range(36)]
    mut = EXTRA["mutacao"]
    return "\n".join([
        "# TEST_REPORT — Lobo Insights Industrial V0.2 (Analytics & Dashboard)",
        f"> {AVISO}",
        "",
        "Relatório **gerado por `scripts/executar_testes_v0_2.py`**: cada PASS/FAIL foi produzido por uma execução real; teste que não pôde rodar aparece como **SKIP** (nunca PASS). "
        "A suíte da V0.1.1 (`scripts/executar_testes.py`, byte a byte igual à da V0.1.1) é executada **inteira e sem alteração** (T01–T103); os testes novos da V0.2 continuam a numeração (T104 em diante). "
        "O relatório histórico da V0.1.1 está preservado em `docs/TEST_REPORT_V0_1_1.md` e o da V0.1 em `docs/TEST_REPORT_V0_1.md`. Para reproduzir: `python scripts/executar_testes_v0_2.py --verificar`.",
        "",
        "## 1. Resumo",
        f"- Testes automatizados: **{len(L)}** — PASS **{npass}**, FAIL **{nfail}**, SKIP **{nskip}**.",
        f"- Suíte legada da V0.1.1 (T01–T{n_legado}): {resumo_legado}",
        f"- Testes novos da V0.2 (T{n_legado + 1}–T{len(L)}): **{len(L) - n_legado}** — PASS **{sum(l[5] == 'PASS' for l in L[n_legado:])}**, FAIL **{sum(l[5] == 'FAIL' for l in L[n_legado:])}**, SKIP **{sum(l[5] == 'SKIP' for l in L[n_legado:])}**.",
        f"- Verificações de qualidade dos CSVs (Python, Q01–Q36): **{len(tab_q)}** — PASS **{int((tab_q['Status'] == 'PASS').sum())}**, FAIL **{int((tab_q['Status'] == 'FAIL').sum())}**; no XLSX V0.2 (aba Qualidade, valores em cache): **{sum(x[2] == 'PASS' for x in q_wb)}/36 PASS**. A comparação do vetor de erros Excel × Python é feita pela suíte legada sobre o XLSX V0.1.1; na V0.2 comprova-se que a aba Qualidade é idêntica à da V0.1.1 (fórmulas e valores; ver 'Preservação da base').",
        f"- Ambiente: Python {platform.python_version()}, pandas {pd.__version__}, openpyxl {openpyxl.__version__}, {lo}, {platform.system()} {platform.machine()}.",
        f"- SHA-256 do XLSX V0.2 (`excel/{NOME_XLSX_V02}`): `{sha256(XLSX02)}`",
        f"- SHA-256 do XLSX histórico V0.1.1 (inalterado): `{sha256(XLSX01)}`",
        "- SHA-256 dos CSVs (inalterados): " + "; ".join(f"`{n}` = `{sha256(DATA / n)}`" for n in base.ARQ),
        "",
        "## 2. TESTADO × NÃO TESTADO",
        "**TESTADO (executado):** toda a suíte da V0.1.1; preservação das abas de dados/Qualidade/Indicadores/Parametros da V0.1.1 (célula a célula); os 6 KPIs (referências do enunciado × XLSX V0.1.1 × Indicadores V0.2 × cartões do DASHBOARD × cálculo independente dos CSVs); "
        "todas as tabelas da aba Analises contra um cálculo independente em Python (Decimal); ausência de erros de fórmula e de valores fixos; ausência de funções de previsão e de funções não portáveis; "
        "estrutura e referências dos 9 gráficos (XML); layout/impressão; hiperlinks internos; reconstrução idêntica do XLSX; **teste de mutação** (dados alterados ⇒ tudo recalcula e continua batendo); "
        "renderização do DASHBOARD em PDF/PNG pelo LibreOffice; documentação.",
        "",
        "**NÃO TESTADO (declarado, não assumido):**",
        "- **Microsoft Excel real** (Windows/Mac/Online): o arquivo foi gerado com openpyxl, recalculado e renderizado com LibreOffice. A aparência e o comportamento dos gráficos no Excel (fontes, espaçamentos, rótulos, eixo secundário do Pareto) **não foram verificados**.",
        "- **Excel em pt-BR / outras localidades:** as fórmulas são gravadas em inglês e o Excel as traduz; os textos de leitura usam `ROUND`/`FIXED` (sem `TEXT` com máscara) para o separador decimal seguir a localidade, mas isso **não foi verificado** fora do LibreOffice em inglês. Os valores em cache (o que aparece antes do recálculo, em visualizadores) usam ponto decimal.",
        "- **Pré-visualizadores que não recalculam** (Quick Look, visualizadores de e-mail/celular): os números vêm do cache gravado, mas os **gráficos podem não aparecer** (o arquivo não grava cache dentro do XML dos gráficos).",
        "- **Filtros/segmentações:** não foram implementados (decisão técnica em `docs/V0_2_NOTES.md`); portanto não há o que testar.",
        "- Execução em Windows/macOS, outras versões de Python/bibliotecas/LibreOffice; instalação limpa por `pip install -r requirements.txt` (sem rede neste ambiente).",
        "- Inspeção visual **humana** de todas as abas: foram inspecionadas as páginas do DASHBOARD (3) e amostras das demais abas (seção 8); o restante não foi visto.",
        "- Tabelas dinâmicas, IA, classificador, Power BI, n8n, GitHub: fora do escopo.",
        "- O conteúdo do ZIP final é conferido após o empacotamento, fora deste relatório (um arquivo não pode conter a prova do próprio ZIP).",
        "",
        "## 3. Verificações de qualidade dos dados (Python, sobre os CSVs)",
        md(["Teste", "Regra", "Resultado", "Quantidade de erros", "Status"], [(f"{r.ID} · {r.Teste}", r.Regra, r.Resultado, r["Quantidade de erros"], r.Status) for _, r in tab_q.iterrows()]),
        "",
        "## 4. Testes automatizados (T01–T103 legados; T104 em diante novos)",
        md(["ID", "Área", "Cenário", "Esperado", "Obtido", "PASS/FAIL"], L),
        "",
        "## 5. Regressão dos 6 KPIs oficiais",
        md(["KPI", "Referência do enunciado", "XLSX V0.1.1", "Indicadores V0.2", "Cartão do DASHBOARD", "CSV independente", "Status"],
           [(n, e, a, b, c, d, "PASS" if ok else "FAIL") for n, e, a, b, c, d, ok in EXTRA["regressao"]]),
        "",
        "## 6. Inventário dos gráficos (lido do XML do XLSX)",
        md(["Arquivo", "Tipo", "Direção", "Séries", "Pontos", "Rótulos", "Referências (todas na aba Analises)"],
           [(g["arquivo"].split("/")[-1], g["tipos"], g["dir"], g["series"], g["pontos"], g["rotulos"], "; ".join(g["refs"])) for g in EXTRA["inventario_graficos"]]),
        "",
        "## 7. Teste de mutação (prova de que nada é fixo)",
        "Uma cópia **em memória** dos dados é alterada (os CSVs oficiais não são tocados): consumo de MAT010 vira 'Ajuste'; MAT019 passa a 300 unidades; MAT004 muda de mês; estoques de MAT001, MAT005, MAT008 e MAT016 mudam. "
        "O XLSX V0.2 é reconstruído e recalculado com esses dados e comparado, tabela a tabela, com um cálculo independente em Python. "
        + ("Amostra dos indicadores que mudam:" if mut else "**NÃO EXECUTADO** (LibreOffice ausente)."),
        "",
        (md(["Indicador", "Dados oficiais", "Esperado com dados alterados", "Obtido no Excel recalculado"], [(a, b, c, d) for a, b, c, d in mut]) if mut else ""),
        "",
        "## 8. Inspeção visual",
        ins,
        "",
        "## 9. Registro de problemas e observações da execução",
        obs,
        "",
    ])


def gerar_previas():
    """Salva as 3 páginas do DASHBOARD (renderizadas pelo LibreOffice) em images/dashboard_v0_2_pagina_N.png."""
    if not (localizar_soffice() and shutil.which("pdftoppm")):
        return []
    with tempfile.TemporaryDirectory() as tmp:
        pdf, _ = renderizar(XLSX02, tmp)
        subprocess.run(["pdftoppm", "-r", "110", "-f", "1", "-l", "3", "-png", str(pdf), str(Path(tmp) / "prev")], capture_output=True, timeout=180)
        (ROOT / "images").mkdir(exist_ok=True)
        saidas = []
        for i, png in enumerate(sorted(Path(tmp).glob("prev-*.png")), start=1):
            destino = ROOT / "images" / f"dashboard_v0_2_pagina_{i}.png"
            shutil.copy(png, destino)
            saidas.append(destino.name)
    return saidas


def main(escrever):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        codigo_legado = base.main(escrever=False)
    saida_legado = buf.getvalue().strip().splitlines()
    n_legado = len(base.LINHAS)
    EXTRA["n_legado"] = n_legado
    resumo = next((l for l in saida_legado if l.startswith("Testes automatizados")), "")
    resumo += " | " + " | ".join(l for l in saida_legado if l.startswith(("Verificações", "Detecção")))
    print("SUÍTE LEGADA V0.1.1:", resumo, f"(código de saída {codigo_legado})")
    dfs, ref, _ = testes_v0_2()
    L = base.LINHAS
    novos = L[n_legado:]
    falhas, pulados = [l for l in L if l[5] == "FAIL"], [l for l in L if l[5] == "SKIP"]
    print(f"Testes novos da V0.2: {len(novos)} | PASS {sum(l[5] == 'PASS' for l in novos)} | FAIL {sum(l[5] == 'FAIL' for l in novos)} | SKIP {sum(l[5] == 'SKIP' for l in novos)}")
    print(f"TOTAL: {len(L)} | PASS {len(L) - len(falhas) - len(pulados)} | FAIL {len(falhas)} | SKIP {len(pulados)}")
    for l in falhas:
        print("FAIL:", l)
    for l in pulados:
        print("SKIP:", l[0], l[2][:90], "->", l[4])
    if escrever:
        wv = openpyxl.load_workbook(XLSX02, data_only=True)
        (DOCS / "TEST_REPORT.md").write_text(relatorio(validar_dados.tabela(validar_dados.executar(dfs)), wv, dfs, resumo, n_legado), encoding="utf-8")
        print("TEST_REPORT.md escrito; prévias:", gerar_previas())
    if falhas or codigo_legado == 1:
        return 1
    return 2 if (pulados or codigo_legado == 2) else 0


if __name__ == "__main__":
    sys.exit(main(escrever="--verificar" not in sys.argv))
