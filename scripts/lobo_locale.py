from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation


_NUMERO_RE = re.compile(
    r"(?<![\w])"
    r"[-+]?"
    r"(?:\d{1,3}(?:[.,]\d{3})+|\d+)"
    r"(?:[.,]\d+)?"
    r"%?"
    r"(?![\w])"
)


def _candidatos_numericos(token: str) -> set[Decimal]:
    """
    Converte uma representação numérica localizada em possíveis valores.

    Exemplos considerados equivalentes:
        9.2       <-> 9,2
        136.70    <-> 136,70
        6,402.20  <-> 6.402,20
        24,606.20 <-> 24.606,20
        1,647     <-> 1.647

    Não torna valores numericamente diferentes equivalentes.
    """

    texto = token.strip()
    percentual = texto.endswith("%")

    if percentual:
        texto = texto[:-1]

    candidatos: set[Decimal] = set()

    def adicionar(valor: str) -> None:
        try:
            numero = Decimal(valor)
            if percentual:
                numero /= Decimal("100")
            candidatos.add(numero)
        except InvalidOperation:
            pass

    tem_ponto = "." in texto
    tem_virgula = "," in texto

    if tem_ponto and tem_virgula:
        # O último separador é tratado como decimal.
        if texto.rfind(".") > texto.rfind(","):
            adicionar(texto.replace(",", ""))
        else:
            adicionar(texto.replace(".", "").replace(",", "."))

        return candidatos

    if tem_ponto or tem_virgula:
        separador = "." if tem_ponto else ","
        partes = texto.split(separador)

        # Interpretação decimal.
        adicionar(texto.replace(",", "."))

        # Interpretação como agrupamento de milhares.
        if (
            len(partes) > 1
            and all(parte.isdigit() or (i == 0 and parte.lstrip("+-").isdigit())
                    for i, parte in enumerate(partes))
            and all(len(parte) == 3 for parte in partes[1:])
        ):
            adicionar("".join(partes))

        return candidatos

    adicionar(texto)
    return candidatos


def numeros_localizados_equivalentes(a: str, b: str) -> bool:
    candidatos_a = _candidatos_numericos(a)
    candidatos_b = _candidatos_numericos(b)

    return bool(candidatos_a and candidatos_b and candidatos_a.intersection(candidatos_b))


def textos_localizados_equivalentes(a: object, b: object) -> bool:
    """
    Compara textos preservando todo o conteúdo não numérico.

    Apenas diferenças de formatação numérica por localidade são aceitas.
    """

    if a == b:
        return True

    if not isinstance(a, str) or not isinstance(b, str):
        return False

    partes_a = _NUMERO_RE.split(a)
    partes_b = _NUMERO_RE.split(b)

    if partes_a != partes_b:
        return False

    numeros_a = _NUMERO_RE.findall(a)
    numeros_b = _NUMERO_RE.findall(b)

    if len(numeros_a) != len(numeros_b):
        return False

    return all(
        numeros_localizados_equivalentes(x, y)
        for x, y in zip(numeros_a, numeros_b)
    )