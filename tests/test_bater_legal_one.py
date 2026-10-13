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
    # Publicação 3 gera 2 linhas (2 CNJs que batem), publicação 1 gera 1 linha.
    assert resultado.linhas_geradas == 3

    from openpyxl import load_workbook

    wb = load_workbook(saida_path)
    ws = wb.active
    linhas = list(ws.iter_rows(values_only=True))
    cabecalho = linhas[0]
    assert cabecalho[-1] == COLUNA_CNJ_SAIDA
    assert cabecalho[:3] == ("Diário", "Data da Publicação", "Conteúdo")

    cnjs_na_saida = {linha[-1] for linha in linhas[1:]}
    assert cnjs_na_saida == {
        "0118284-78.2026.8.04.1000",
        "0667841-11.2025.8.04.1000",
    }
    # Processo sem match na base não deve aparecer na saída.
    assert all("0000000-00.2020.8.04.0000" not in str(linha) for linha in linhas[1:])
