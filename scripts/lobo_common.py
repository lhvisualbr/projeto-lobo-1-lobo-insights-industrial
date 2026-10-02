# -*- coding: utf-8 -*-
"""Definições compartilhadas do Lobo Insights Industrial V0.1.
Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio."""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
EXCEL = ROOT / "excel"
DOCS = ROOT / "docs"

SEP = ";"                      # separador de campos dos CSVs (conforme o PDF)
ENC = "utf-8-sig"              # UTF-8 com BOM (o Excel abre os acentos corretamente)
SEED = 20260601                # semente da geração -> CSVs reproduzíveis

CATEGORIAS = ["Abrasivos", "Soldagem", "Corte", "Elétrica", "EPI", "Ferramentas"]
UNIDADES = ["UN", "CX", "PCT", "PAR"]
CRITICIDADES = ["Baixa", "Média", "Alta"]
CENTROS = ["MANUTENCAO", "SOLDAGEM", "FABRICACAO", "MECANICA", "ELETRICA"]
TIPOS_MOV = ["Consumo", "Ajuste"]

PERIODO_INICIO = "2026-06-01"
PERIODO_FIM = "2026-08-31"
MARGEM_ATENCAO = 0.25          # ATENÇÃO: Atual > Mínimo e Atual <= Mínimo x (1 + margem)

COLS_MATERIAIS = ["Material_ID", "Material", "Categoria", "Unidade",
                  "Custo_Unitario", "Criticidade", "Fornecedor"]
COLS_ESTOQUE = ["Material_ID", "Material", "Categoria", "Unidade", "Estoque_Atual",
                "Estoque_Minimo", "Custo_Unitario", "Lead_Time_Dias", "Fornecedor",
                "Criticidade"]
COLS_CONSUMO = ["Movimento_ID", "Data", "Material_ID", "Material", "Categoria",
                "Quantidade", "Unidade", "Custo_Unitario", "Centro_Trabalho",
                "Tipo_Movimentacao"]

# (ID, Material, Categoria, Unidade, Custo_Unitario, Criticidade, Fornecedor)
MATERIAIS = [
    ("MAT001", "Disco de Desbaste 7 pol",      "Abrasivos",   "UN",   9.80, "Média", "Fornecedor Beta"),
    ("MAT002", "Disco Flap 4,5 pol Grão 60",   "Abrasivos",   "UN",   7.40, "Baixa", "Fornecedor Beta"),
    ("MAT003", "Lixa de Ferro Grão 80",        "Abrasivos",   "PCT", 38.00, "Baixa", "Fornecedor Gama"),
    ("MAT004", "Bico de Contato MIG 1,0 mm",   "Soldagem",    "PCT", 45.00, "Média", "Fornecedor Delta"),
    ("MAT005", "Eletrodo Revestido 2,5 mm",    "Soldagem",    "CX",  96.00, "Alta",  "Fornecedor Alfa"),
    ("MAT006", "Arame MIG 1,0 mm Bobina",      "Soldagem",    "UN", 210.00, "Alta",  "Fornecedor Delta"),
    ("MAT007", "Disco de Corte 7 pol",         "Corte",       "UN",  12.50, "Alta",  "Fornecedor Alfa"),
    ("MAT008", "Disco de Corte 4,5 pol",       "Corte",       "UN",   5.90, "Média", "Fornecedor Alfa"),
    ("MAT009", "Serra Copo 51 mm",             "Corte",       "UN",  46.00, "Baixa", "Fornecedor Sigma"),
    ("MAT010", "Lâmina de Serra Sabre",        "Corte",       "PCT", 62.00, "Média", "Fornecedor Sigma"),
    ("MAT011", "Cabo Flexível 2,5 mm² Rolo",   "Elétrica",    "UN", 285.00, "Média", "Fornecedor Beta"),
    ("MAT012", "Fita Isolante 19 mm",          "Elétrica",    "PCT", 42.00, "Baixa", "Fornecedor Gama"),
    ("MAT013", "Terminal de Compressão 6 mm²", "Elétrica",    "PCT", 28.00, "Média", "Fornecedor Delta"),
    ("MAT014", "Disjuntor Tripolar 32 A",      "Elétrica",    "UN", 118.00, "Alta",  "Fornecedor Delta"),
    ("MAT015", "Luva de Raspa",                "EPI",         "PAR", 15.00, "Média", "Fornecedor Sigma"),
    ("MAT016", "Óculos de Proteção Incolor",   "EPI",         "UN",   6.50, "Média", "Fornecedor Sigma"),
    ("MAT017", "Máscara Respiratória PFF2",    "EPI",         "CX",  72.00, "Alta",  "Fornecedor Sigma"),
    ("MAT018", "Chave Combinada 17 mm",        "Ferramentas", "UN",  24.00, "Baixa", "Fornecedor Gama"),
    ("MAT019", "Trena de Aço 5 m",             "Ferramentas", "UN",  19.00, "Baixa", "Fornecedor Beta"),
    ("MAT020", "Broca de Aço Rápido 10 mm",    "Ferramentas", "UN",  21.00, "Média", "Fornecedor Alfa"),
]

# Material_ID -> (Estoque_Atual, Estoque_Minimo, Lead_Time_Dias)
ESTOQUE = {
    "MAT001": (120, 40, 7),  "MAT002": (30, 30, 7),   "MAT003": (34, 10, 10),
    "MAT004": (9, 8, 14),    "MAT005": (7, 12, 12),   "MAT006": (14, 6, 21),
    "MAT007": (74, 60, 7),   "MAT008": (210, 80, 7),  "MAT009": (12, 4, 15),
    "MAT010": (5, 6, 18),    "MAT011": (8, 3, 20),    "MAT012": (25, 10, 10),
    "MAT013": (14, 12, 14),  "MAT014": (3, 4, 21),    "MAT015": (22, 30, 9),
    "MAT016": (90, 40, 5),   "MAT017": (11, 10, 12),  "MAT018": (12, 4, 10),
    "MAT019": (9, 3, 8),     "MAT020": (26, 10, 12),
}


def status_estoque(atual, minimo, margem=MARGEM_ATENCAO):
    """Regra transparente de status (NÃO é previsão): estado atual versus mínimo."""
    if atual <= minimo:
        return "REPOR"
    if atual <= minimo * (1 + margem):
        return "ATENÇÃO"
    return "OK"


# ---------------------------------------------------------------------------
# Especificação das verificações de qualidade.
# Python (validar_dados.py) e Excel (aba Qualidade) usam o MESMO ID de teste.
# ---------------------------------------------------------------------------
M, E, C = "tbMateriais", "tbEstoque", "tbConsumo"


def _dup(t, c):
    return f"=SUMPRODUCT(--(COUNTIF({t}[{c}],{t}[{c}])>1))"


def _num(t, c, cond):
    # Se houver texto/vazio na coluna numérica, conta essas células e ignora o resto.
    return (f"=IF(ROWS({t}[{c}])-COUNT({t}[{c}])>0,"
            f"ROWS({t}[{c}])-COUNT({t}[{c}]),{cond})")


def _miss(t):
    return f"=SUMPRODUCT(--(COUNTIF({M}[Material_ID],{t}[Material_ID])=0))"


def _mis(t, f):
    return (f"=SUMPRODUCT((COUNTIF({M}[Material_ID],{t}[Material_ID])>0)*"
            f"(COUNTIFS({M}[Material_ID],{t}[Material_ID],{M}[{f}],{t}[{f}])=0))")


def _dom(lst, t, c):
    return f"=SUMPRODUCT(--(COUNTIF({lst},{t}[{c}])=0))"


CHECKS = [
    ("Q01", "IDs duplicados – Materiais", "Material_ID não se repete (conta linhas envolvidas em duplicidade)", _dup(M, "Material_ID")),
    ("Q02", "IDs duplicados – Estoque", "Material_ID não se repete na posição de estoque", _dup(E, "Material_ID")),
    ("Q03", "IDs duplicados – Consumo", "Movimento_ID não se repete", _dup(C, "Movimento_ID")),
    ("Q04", "Nulos indevidos – Materiais", "Nenhuma célula vazia na tabela", f"=COUNTBLANK({M})"),
    ("Q05", "Nulos indevidos – Estoque", "Nenhuma célula vazia na tabela", f"=COUNTBLANK({E})"),
    ("Q06", "Nulos indevidos – Consumo", "Nenhuma célula vazia na tabela", f"=COUNTBLANK({C})"),
    ("Q07", "Quantidade inválida – Consumo", "Quantidade numérica, inteira e > 0",
     _num(C, "Quantidade", f'COUNTIF({C}[Quantidade],"<=0")+SUMPRODUCT(--({C}[Quantidade]<>INT({C}[Quantidade])))')),
    ("Q08", "Custo <= 0 – Materiais", "Custo_Unitario numérico e > 0", _num(M, "Custo_Unitario", f'COUNTIF({M}[Custo_Unitario],"<=0")')),
    ("Q09", "Custo <= 0 – Estoque", "Custo_Unitario numérico e > 0", _num(E, "Custo_Unitario", f'COUNTIF({E}[Custo_Unitario],"<=0")')),
    ("Q10", "Custo <= 0 – Consumo", "Custo_Unitario numérico e > 0", _num(C, "Custo_Unitario", f'COUNTIF({C}[Custo_Unitario],"<=0")')),
    ("Q11", "Material inexistente – Estoque", "Material_ID do estoque existe no catálogo", _miss(E)),
    ("Q12", "Material inexistente – Consumo", "Material_ID do consumo existe no catálogo", _miss(C)),
    ("Q13", "Catálogo sem estoque", "Todo material do catálogo tem posição de estoque",
     f"=SUMPRODUCT(--(COUNTIF({E}[Material_ID],{M}[Material_ID])=0))"),
    ("Q14", "Nome divergente – Estoque", "Material do estoque = catálogo", _mis(E, "Material")),
    ("Q15", "Nome divergente – Consumo", "Material do consumo = catálogo", _mis(C, "Material")),
    ("Q16", "Categoria divergente – Estoque", "Categoria do estoque = catálogo", _mis(E, "Categoria")),
    ("Q17", "Categoria divergente – Consumo", "Categoria do consumo = catálogo", _mis(C, "Categoria")),
    ("Q18", "Unidade divergente – Estoque", "Unidade do estoque = catálogo", _mis(E, "Unidade")),
    ("Q19", "Unidade divergente – Consumo", "Unidade do consumo = catálogo", _mis(C, "Unidade")),
    ("Q20", "Custo divergente – Estoque", "Custo_Unitario do estoque = catálogo", _mis(E, "Custo_Unitario")),
    ("Q21", "Custo divergente – Consumo", "Custo_Unitario do consumo = catálogo", _mis(C, "Custo_Unitario")),
    ("Q22", "Criticidade divergente – Estoque", "Criticidade do estoque = catálogo", _mis(E, "Criticidade")),
    ("Q23", "Fornecedor divergente – Estoque", "Fornecedor do estoque = catálogo", _mis(E, "Fornecedor")),
    ("Q24", "Datas inválidas – Consumo", "Data válida e dentro do período definido na aba Parametros",
     f'=SUMPRODUCT(--NOT(ISNUMBER({C}[Data])))+COUNTIF({C}[Data],"<"&pInicioPeriodo)+COUNTIF({C}[Data],">"&pFimPeriodo)'),
    ("Q25", "Cobertura dos 3 meses – Consumo", "Cada mês do período tem >= 1 movimentação (conta meses sem registro)",
     f"=SUMPRODUCT(--(COUNTIF({C}[Mes],lstMeses)=0))"),
    ("Q26", "Estoque negativo", "Estoque_Atual numérico e >= 0", _num(E, "Estoque_Atual", f'COUNTIF({E}[Estoque_Atual],"<0")')),
    ("Q27", "Estoque mínimo negativo", "Estoque_Minimo numérico e >= 0", _num(E, "Estoque_Minimo", f'COUNTIF({E}[Estoque_Minimo],"<0")')),
    ("Q28", "Lead time inválido", "Lead_Time_Dias numérico, inteiro e > 0",
     _num(E, "Lead_Time_Dias", f'COUNTIF({E}[Lead_Time_Dias],"<=0")+SUMPRODUCT(--({E}[Lead_Time_Dias]<>INT({E}[Lead_Time_Dias])))')),
    ("Q29", "Categoria inválida – Materiais", "Categoria pertence à lista permitida", _dom("lstCategorias", M, "Categoria")),
    ("Q30", "Categoria inválida – Estoque", "Categoria pertence à lista permitida", _dom("lstCategorias", E, "Categoria")),
    ("Q31", "Categoria inválida – Consumo", "Categoria pertence à lista permitida", _dom("lstCategorias", C, "Categoria")),
    ("Q32", "Unidade inválida – Materiais", "Unidade em UN, CX, PCT ou PAR", _dom("lstUnidades", M, "Unidade")),
    ("Q33", "Criticidade inválida – Materiais", "Criticidade em Baixa, Média ou Alta", _dom("lstCriticidade", M, "Criticidade")),
    ("Q34", "Centro inválido – Consumo", "Centro_Trabalho na lista permitida", _dom("lstCentros", C, "Centro_Trabalho")),
    ("Q35", "Tipo de movimentação inválido", "Tipo_Movimentacao em Consumo ou Ajuste", _dom("lstTipoMov", C, "Tipo_Movimentacao")),
    ("Q36", "Fornecedor fora do padrão fictício", "Fornecedor começa com 'Fornecedor ' (nome fictício)",
     f'=SUMPRODUCT(--(LEFT({M}[Fornecedor],11)<>"Fornecedor "))'),
]
