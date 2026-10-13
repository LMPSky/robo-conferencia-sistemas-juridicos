import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from robo_conferencia.cnj import find_cnj_numbers, format_cnj, normalize_cnj  # noqa: E402


def test_normalize_cnj_remove_pontuacao():
    assert normalize_cnj("0118284-78.2026.8.04.1000") == "01182847820268041000"[0:20]


def test_format_cnj_a_partir_de_digitos():
    assert format_cnj("01182847820268041000") == "0118284-78.2026.8.04.1000"


def test_find_cnj_numbers_em_texto_livre():
    texto = (
        "Processo nº 0118284-78.2026.8.04.1000 publicado. "
        "Ver também o processo 0667841-11.2025.8.04.1000 relacionado."
    )
    assert find_cnj_numbers(texto) == [
        "0118284-78.2026.8.04.1000",
        "0667841-11.2025.8.04.1000",
    ]


def test_find_cnj_numbers_sem_pontuacao():
    texto = "Processo 01182847820268041000 sem pontuacao"
    assert find_cnj_numbers(texto) == ["0118284-78.2026.8.04.1000"]


def test_find_cnj_numbers_ignora_duplicatas():
    texto = (
        "0118284-78.2026.8.04.1000 ... 01182847820268041000 "
        "(mesmo número, formatos diferentes)"
    )
    assert find_cnj_numbers(texto) == ["0118284-78.2026.8.04.1000"]


def test_find_cnj_numbers_texto_vazio():
    assert find_cnj_numbers("") == []
    assert find_cnj_numbers(None) == []
