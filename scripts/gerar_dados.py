# -*- coding: utf-8 -*-
"""Gera os três CSVs sintéticos da V0.1 (determinístico: mesma SEED => mesmos arquivos).

Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.
Nenhuma conclusão é forçada: a distribuição vem de pesos e sorteios simples;
os insights devem surgir da análise posterior.
"""
import csv
import random
from datetime import date, timedelta

from lobo_common import (CENTROS, COLS_CONSUMO, COLS_ESTOQUE, COLS_MATERIAIS, DATA,
                         ENC, ESTOQUE, MATERIAIS, SEED, SEP)

N_MOVIMENTOS = 180
INICIO, FIM = date(2026, 6, 1), date(2026, 8, 31)

# Peso relativo de frequência e faixa de quantidade por movimentação (lo, hi)
PERFIL = {
    "MAT001": (9, 5, 30),  "MAT002": (9, 5, 30),  "MAT003": (5, 1, 4),   "MAT004": (5, 1, 3),
    "MAT005": (8, 1, 4),   "MAT006": (4, 1, 2),   "MAT007": (12, 5, 40), "MAT008": (12, 10, 60),
    "MAT009": (2, 1, 3),   "MAT010": (3, 1, 2),   "MAT011": (2, 1, 2),   "MAT012": (3, 1, 3),
    "MAT013": (4, 1, 3),   "MAT014": (1.5, 1, 2), "MAT015": (10, 2, 12), "MAT016": (8, 2, 10),
    "MAT017": (5, 1, 3),   "MAT018": (1.5, 1, 2), "MAT019": (1, 1, 2),   "MAT020": (3, 1, 4),
}
# Pesos por centro, na ordem de CENTROS: MANUTENCAO, SOLDAGEM, FABRICACAO, MECANICA, ELETRICA
CENTRO_POR_CATEGORIA = {
    "Abrasivos":   [0.25, 0.20, 0.35, 0.18, 0.02],
    "Soldagem":    [0.15, 0.55, 0.25, 0.05, 0.00],
    "Corte":       [0.22, 0.15, 0.38, 0.24, 0.01],
    "Elétrica":    [0.35, 0.01, 0.04, 0.05, 0.55],
    "EPI":         [0.30, 0.20, 0.20, 0.15, 0.15],
    "Ferramentas": [0.35, 0.10, 0.15, 0.30, 0.10],
}


def _dias_com_peso(rng):
    """Peso por dia: útil=1, sábado=0.15, domingo=0; ruído por dia e fator por mês."""
    fator_mes = {m: rng.uniform(0.85, 1.15) for m in (6, 7, 8)}
    dias, pesos = [], []
    d = INICIO
    while d <= FIM:
        base = 0.0 if d.weekday() == 6 else (0.15 if d.weekday() == 5 else 1.0)
        dias.append(d)
        pesos.append(base * rng.uniform(0.4, 1.8) * fator_mes[d.month])
        d += timedelta(days=1)
    return dias, pesos


def gerar():
    rng = random.Random(SEED)
    mats = {m[0]: m for m in MATERIAIS}
    ids = [m[0] for m in MATERIAIS]
    pesos_mat = [PERFIL[i][0] for i in ids]
    dias, pesos_dia = _dias_com_peso(rng)

    movs = []
    for _ in range(N_MOVIMENTOS):
        dia = rng.choices(dias, weights=pesos_dia, k=1)[0]
        mid = rng.choices(ids, weights=pesos_mat, k=1)[0]
        _, mat, cat, uni, custo, _crit, _forn = mats[mid]
        centro = rng.choices(CENTROS, weights=CENTRO_POR_CATEGORIA[cat], k=1)[0]
        _, lo, hi = PERFIL[mid]
        qtd = min(hi, lo + int((hi - lo + 1) * rng.random() ** 1.6))
        movs.append([dia, mid, mat, cat, qtd, uni, custo, centro, "Consumo"])
    movs.sort(key=lambda r: r[0])           # ordena por data (sort estável)

    consumo = []
    for n, (dia, mid, mat, cat, qtd, uni, custo, centro, tipo) in enumerate(movs, start=1):
        consumo.append([f"MOV{n:04d}", dia.isoformat(), mid, mat, cat, qtd, uni,
                        f"{custo:.2f}", centro, tipo])

    materiais = [[i, n, c, u, f"{p:.2f}", k, f] for (i, n, c, u, p, k, f) in MATERIAIS]
    estoque = []
    for (i, n, c, u, p, k, f) in MATERIAIS:
        atual, minimo, lead = ESTOQUE[i]
        estoque.append([i, n, c, u, atual, minimo, f"{p:.2f}", lead, f, k])
    return materiais, estoque, consumo


def _salvar(nome, cabecalho, linhas, pasta=DATA):
    with open(pasta / nome, "w", newline="", encoding=ENC) as fh:
        w = csv.writer(fh, delimiter=SEP, lineterminator="\n")
        w.writerow(cabecalho)
        w.writerows(linhas)


def main(pasta=DATA):
    pasta.mkdir(exist_ok=True)
    materiais, estoque, consumo = gerar()
    _salvar("materiais_ficticios.csv", COLS_MATERIAIS, materiais, pasta)
    _salvar("estoque_ficticio.csv", COLS_ESTOQUE, estoque, pasta)
    _salvar("consumo_ficticio.csv", COLS_CONSUMO, consumo, pasta)
    print(f"materiais={len(materiais)} estoque={len(estoque)} consumo={len(consumo)}")


if __name__ == "__main__":
    main()
