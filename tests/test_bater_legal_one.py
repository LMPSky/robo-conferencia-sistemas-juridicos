import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from openpyxl import Workbook  # noqa: E402

from robo_conferencia.bater_legal_one import (  # noqa: E402
    COLUNA_CNJ_SAIDA,
    bater_publicacoes_legal_one,
)


def _criar_publicacoes(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["Diário", "Data da Publicação", "Conteúdo"])
    ws.append(
        [
            "Diário X",
            "2026-01-01",
            "Processo 0118284-78.2026.8.04.1000 publicado hoje.",
        ]
    )
    ws.append(
        [
            "Diário X",
            "2026-01-02",
            "Processo 0000000-00.2020.8.04.0000 sem match na base.",
        ]
    )
    ws.append(
        [
            "Diário X",
            "2026-01-03",
            (
                "Dois processos: 0118284-78.2026.8.04.1000 e "
                "0667841-11.2025.8.04.1000."
            ),
        ]
    )
    wb.save(path)


def _criar_legal_one(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["Número do Processo", "Cliente"])
    ws.append(["0118284-78.2026.8.04.1000", "Cliente A"])
    ws.append(["0667841-11.2025.8.04.1000", "Cliente B"])
    wb.save(path)


def test_bater_publicacoes_legal_one(tmp_path: Path):
    pub_path = tmp_path / "publicacoes.xlsx"
    legal_one_path = tmp_path / "legal_one.xlsx"
    saida_path = tmp_path / "saida.xlsx"

    _criar_publicacoes(pub_path)
    _criar_legal_one(legal_one_path)

    resultado = bater_publicacoes_legal_one(pub_path, legal_one_path, saida_path)

    assert resultado.processos_legal_one == 2
    assert resultado.publicacoes_lidas == 3
    # Apenas as publicações 1 e 3 têm CNJ que bate com a base do Legal One.
    assert resultado.publicacoes_com_match == 2
    # Publicação 3 tem 2 CNJs que batem na mesma linha -> 2 colunas de CNJ.
    assert resultado.max_cnj_por_publicacao == 2

    from openpyxl import load_workbook

    wb = load_workbook(saida_path)
    ws = wb.active
    linhas = list(ws.iter_rows(values_only=True))
    cabecalho = linhas[0]
    assert cabecalho[-2:] == (f"{COLUNA_CNJ_SAIDA} 1", f"{COLUNA_CNJ_SAIDA} 2")
    assert cabecalho[:3] == ("Diário", "Data da Publicação", "Conteúdo")

    # Uma linha de saída por publicação com match (não uma por CNJ).
    assert len(linhas) - 1 == 2

    cnjs_na_saida = set()
    for linha in linhas[1:]:
        cnjs_na_saida.update(v for v in linha[-2:] if v)
    assert cnjs_na_saida == {
        "0118284-78.2026.8.04.1000",
        "0667841-11.2025.8.04.1000",
    }
    # Processo sem match na base não deve aparecer na saída.
    assert all("0000000-00.2020.8.04.0000" not in str(linha) for linha in linhas[1:])

    # A publicação 3 (dois processos) deve ter ambas as colunas preenchidas,
    # e a publicação 1 (um processo) deve ter a segunda coluna vazia.
    por_data = {linha[1]: linha for linha in linhas[1:]}
    assert por_data["2026-01-01"][-2:] == ("0118284-78.2026.8.04.1000", None)
    assert por_data["2026-01-03"][-2:] == (
        "0118284-78.2026.8.04.1000",
        "0667841-11.2025.8.04.1000",
    )
