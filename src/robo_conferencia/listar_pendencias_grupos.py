"""Triagem de um export "em lote" do GED que só traz nomes de grupos/pastas.

Contexto de negócio
--------------------
Alguns sistemas de GED permitem baixar todos os itens de uma área de uma vez,
mas quando os processos estão organizados em pastas/grupos dentro do sistema,
o resultado desse download em massa é uma única planilha com apenas o *nome*
de cada grupo (não o conteúdo de cada processo dentro dele). Abrir cada grupo
individualmente para descobrir quais processos ele contém é inviável quando
há centenas de grupos.

Este módulo reduz drasticamente esse trabalho manual: se o nome do grupo já
contém (mesmo que em meio a outro texto) um número de processo no padrão CNJ,
o robô extrai esse número automaticamente e já classifica o grupo comparando
com a base do Legal One, sem precisar abrir o grupo. Só restam para
verificação manual os grupos cujo nome não tem nenhum CNJ identificável.

Fluxo
-----
1. Lê a planilha com a lista de grupos (export em lote do GED) e, de uma
   coluna de texto livre, extrai todos os números CNJ encontrados no nome de
   cada grupo.
2. Lê a base de processos do Legal One e carrega o conjunto de CNJs já
   cadastrados.
3. Para cada grupo:
   - se nenhum CNJ foi encontrado no nome, o status é
     ``SEM NÚMERO IDENTIFICADO`` (requer abrir o grupo manualmente);
   - se algum CNJ foi encontrado, gera uma linha por CNJ encontrado, com
     status ``CADASTRADO`` (já existe na base do Legal One) ou ``PENDENTE``
     (não encontrado na base).
4. Gera uma planilha de saída com todas as colunas originais do grupo + o CNJ
   encontrado (quando houver) + o status.
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Sequence

from openpyxl import Workbook

from .bater_legal_one import (
    _carregar_cnj_legal_one,
    _carregar_planilha,
    _resolver_indice_coluna,
)
from .cnj import find_cnj_numbers, normalize_cnj

COLUNA_CNJ_SAIDA = "Número CNJ"
COLUNA_STATUS_SAIDA = "Status"

STATUS_CADASTRADO = "CADASTRADO"
STATUS_PENDENTE = "PENDENTE"
STATUS_SEM_NUMERO = "SEM NÚMERO IDENTIFICADO"


@dataclass
class ResultadoConferenciaGrupos:
    grupos_lidos: int
    grupos_cadastrados: int
    grupos_pendentes: int
    grupos_sem_numero: int
    linhas_geradas: int
    processos_legal_one: int


def conferir_grupos_ged(
    grupos_path: Path,
    legal_one_path: Path,
    saida_path: Path,
    *,
    aba_grupos: Optional[str] = None,
    coluna_texto: str = "A",
    aba_legal_one: Optional[str] = None,
    coluna_cnj_legal_one: Optional[str] = None,
) -> ResultadoConferenciaGrupos:
    """Confere cada grupo listado em `grupos_path` contra a base do Legal One.

    Retorna um `ResultadoConferenciaGrupos` com estatísticas da execução e
    grava a planilha de saída em `saida_path`.
    """
    processos_legal_one = _carregar_cnj_legal_one(
        legal_one_path, aba_legal_one, coluna_cnj_legal_one
    )

    ws_grupos = _carregar_planilha(grupos_path, aba_grupos)
    linhas = list(ws_grupos.iter_rows(values_only=True))
    if not linhas:
        raise ValueError("A planilha de grupos está vazia.")

    cabecalho = list(linhas[0])
    idx_texto = _resolver_indice_coluna(cabecalho, coluna_texto, "planilha de grupos")
    linhas_dados = linhas[1:]

    wb_saida = Workbook()
    ws_saida = wb_saida.active
    ws_saida.title = "Conferência de grupos"
    ws_saida.append([*cabecalho, COLUNA_CNJ_SAIDA, COLUNA_STATUS_SAIDA])

    grupos_cadastrados = 0
    grupos_pendentes = 0
    grupos_sem_numero = 0
    linhas_geradas = 0

    for linha in linhas_dados:
        texto = linha[idx_texto] if idx_texto < len(linha) else None
        numeros_encontrados = find_cnj_numbers(str(texto)) if texto else []

        if not numeros_encontrados:
            grupos_sem_numero += 1
            ws_saida.append([*linha, "", STATUS_SEM_NUMERO])
            linhas_geradas += 1
            continue

        grupo_tem_pendente = False
        grupo_tem_cadastrado = False
        for numero in numeros_encontrados:
            if normalize_cnj(numero) in processos_legal_one:
                status = STATUS_CADASTRADO
                grupo_tem_cadastrado = True
            else:
                status = STATUS_PENDENTE
                grupo_tem_pendente = True
            ws_saida.append([*linha, numero, status])
            linhas_geradas += 1

        # Para o resumo, cada grupo conta em pendente se tiver ao menos um
        # CNJ pendente (prioriza o que precisa de ação), senão em cadastrado.
        if grupo_tem_pendente:
            grupos_pendentes += 1
        elif grupo_tem_cadastrado:
            grupos_cadastrados += 1

    saida_path.parent.mkdir(parents=True, exist_ok=True)
    wb_saida.save(saida_path)

    return ResultadoConferenciaGrupos(
        grupos_lidos=len(linhas_dados),
        grupos_cadastrados=grupos_cadastrados,
        grupos_pendentes=grupos_pendentes,
        grupos_sem_numero=grupos_sem_numero,
        linhas_geradas=linhas_geradas,
        processos_legal_one=len(processos_legal_one),
    )


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Confere uma lista de grupos/pastas do GED (export em lote, só "
            "com o nome de cada grupo) contra a base do Legal One, extraindo "
            "o número CNJ do nome do grupo quando possível. Grupos sem CNJ "
            "identificável no nome precisam ser abertos manualmente."
        )
    )
    parser.add_argument(
        "grupos", type=Path, help="Planilha com a lista de grupos do GED (.xlsx)"
    )
    parser.add_argument("legal_one", type=Path, help="Planilha de base do Legal One (.xlsx)")
    parser.add_argument("saida", type=Path, help="Caminho da planilha de saída (.xlsx)")
    parser.add_argument(
        "--aba-grupos", default=None, help="Nome da aba da planilha de grupos"
    )
    parser.add_argument(
        "--coluna-texto",
        default="A",
        help=(
            "Coluna com o nome/texto de cada grupo. Aceita letra do Excel "
            "(ex: A), índice 0-based ou nome do cabeçalho. Padrão: A"
        ),
    )
    parser.add_argument(
        "--aba-legal-one", default=None, help="Nome da aba da planilha de base do Legal One"
    )
    parser.add_argument(
        "--coluna-cnj-legal-one",
        default=None,
        help=(
            "Coluna com o número CNJ na base do Legal One. Se omitida, é "
            "detectada automaticamente."
        ),
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = _build_arg_parser()
    args = parser.parse_args(argv)

    resultado = conferir_grupos_ged(
        args.grupos,
        args.legal_one,
        args.saida,
        aba_grupos=args.aba_grupos,
        coluna_texto=args.coluna_texto,
        aba_legal_one=args.aba_legal_one,
        coluna_cnj_legal_one=args.coluna_cnj_legal_one,
    )

    print(f"Processos na base do Legal One: {resultado.processos_legal_one}")
    print(f"Grupos lidos: {resultado.grupos_lidos}")
    print(f"Grupos cadastrados: {resultado.grupos_cadastrados}")
    print(f"Grupos pendentes: {resultado.grupos_pendentes}")
    print(
        f"Grupos SEM NÚMERO IDENTIFICADO (abrir manualmente): "
        f"{resultado.grupos_sem_numero}"
    )
    print(f"Linhas geradas na planilha de saída: {resultado.linhas_geradas}")
    print(f"Planilha gerada em: {args.saida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
