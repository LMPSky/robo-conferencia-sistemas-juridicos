"""Extração de números de processo no padrão CNJ a partir de textos livres.

O padrão CNJ (Resolução 65/2008 do CNJ) tem o formato:

    NNNNNNN-DD.AAAA.J.TR.OOOO

- NNNNNNN: número sequencial do processo (7 dígitos)
- DD: dígito verificador (2 dígitos)
- AAAA: ano do ajuizamento (4 dígitos)
- J: órgão do judiciário (1 dígito)
- TR: tribunal (2 dígitos)
- OOOO: unidade de origem (4 dígitos)

O separador pode variar (ou até estar ausente) dependendo de como o texto foi
copiado/extraído de um diário oficial, então o regex aceita pontuação opcional
entre os grupos e a normalização remove toda a pontuação para comparação.
"""
from __future__ import annotations

import re
from typing import Iterable, Iterator, List

# Sequência de 20 dígitos com pontuação opcional nos separadores padrão.
_CNJ_PATTERN = re.compile(
    r"\d{7}\s*-?\s*\d{2}\s*\.?\s*\d{4}\s*\.?\s*\d\s*\.?\s*\d{2}\s*\.?\s*\d{4}"
)

_DIGITS_ONLY = re.compile(r"\D+")


def normalize_cnj(numero: str) -> str:
    """Remove toda a formatação, retornando apenas os 20 dígitos do número CNJ."""
    return _DIGITS_ONLY.sub("", numero)


def format_cnj(numero: str) -> str:
    """Formata uma string de 20 dígitos no padrão NNNNNNN-DD.AAAA.J.TR.OOOO.

    Se `numero` não tiver exatamente 20 dígitos após a normalização, retorna o
    valor normalizado sem formatação (fallback seguro).
    """
    digitos = normalize_cnj(numero)
    if len(digitos) != 20:
        return digitos
    return (
        f"{digitos[0:7]}-{digitos[7:9]}.{digitos[9:13]}."
        f"{digitos[13:14]}.{digitos[14:16]}.{digitos[16:20]}"
    )


def find_cnj_numbers(texto: str) -> List[str]:
    """Retorna, em ordem e sem duplicatas, todos os números CNJ formatados
    encontrados no texto informado.

    A deduplicação é feita pelo valor normalizado (apenas dígitos), mas o
    retorno é sempre formatado em NNNNNNN-DD.AAAA.J.TR.OOOO.
    """
    if not texto:
        return []

    encontrados: List[str] = []
    vistos = set()
    for match in _CNJ_PATTERN.finditer(str(texto)):
        digitos = normalize_cnj(match.group())
        if len(digitos) != 20 or digitos in vistos:
            continue
        vistos.add(digitos)
        encontrados.append(format_cnj(digitos))
    return encontrados


def iter_cnj_numbers(textos: Iterable[str]) -> Iterator[str]:
    """Itera números CNJ (normalizados) encontrados em uma coleção de textos."""
    for texto in textos:
        for numero in find_cnj_numbers(texto):
            yield numero
