"""Testes do batimento entre planilhas de cadastro e a base do Legal One.

As fixtures são geradas programaticamente (nenhum dado real do usuário é
usado), montando a mesma estrutura de pastas esperada pelo robô
(``entrada/base_legal_one``, ``entrada/planilhas_cadastro``,
``entrada/export_ged``) dentro de um diretório temporário.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pandas as pd  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402

from robo_conferencia.batimento_cadastro import (  # noqa: E402
    STATUS_CADASTRADO,
    STATUS_PENDENTE,
    STATUS_SEM_NUMERO,
    executar_batimento,
)


def _criar_base_legal_one(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    # Primeira linha: título solto que deve ser ignorado.
    ws.append(["usuario@empresa.com.br - 01/01/2026"])
    ws.append(
        [
            "Pasta",
            "Número de CNJ",
            "Número antigo",
            "Outro número",
            "Envolvidos",
            "Envolvidos / Posição",
            "Envolvidos / Situação",
        ]
    )
    ws.append(
        [
            "Proc - 0050859/001",
            "0010903-71.2018.5.03.0029",
            "",
            "",
            "Cliente A",
            "Autor",
            "Ativo",
        ]
    )
    ws.append(
        [
            "Proc - 0050900/001",
            "",
            "",
            "2722/2026",
            "Cliente B",
            "Réu",
            "Ativo",
        ]
    )
    ws.append(
        [
            "Proc - 0050950/001",
            "",
            "",
            "26.10.0846.001.00004-3",
            "Cliente C",
            "Autor",
            "Ativo",
        ]
    )
    wb.save(path)


def _criar_planilha_cadastro_advogado(path: Path) -> None:
    """Planilha com cabeçalhos "padrão" e aba chamada CADASTRO."""
    wb = Workbook()
    ws = wb.active
    ws.title = "CADASTRO"
    ws.append(
        [
            "Número CNJ",
            "Outro número",
            "Cliente principal",
            "Responsável principal",
            "Status",
        ]
    )
    # Bate com a base (mesmo CNJ) -> CADASTRADO
    ws.append(
        [
            "0010903-71.2018.5.03.0029",
            "",
            "Cliente A",
            "Dr. Fulano",
            "Novo",
        ]
    )
    # Não bate com nada -> PENDENTE
    ws.append(
        [
            "0099999-99.2026.8.26.0100",
            "",
            "Cliente X",
            "Dra. Sicrana",
            "Novo",
        ]
    )
    # Sem CNJ e sem outro número -> SEM NÚMERO IDENTIFICADO
    ws.append(
        [
            "",
            "",
            "Cliente Y",
            "Dr. Beltrano",
            "Novo",
        ]
    )
    wb.save(path)


def _criar_planilha_cadastro_robo_ti(path: Path) -> None:
    """Planilha com cabeçalhos em variação (acentos/maiúsculas diferentes) e
    sem aba chamada CADASTRO (deve cair no fallback de detectar pela coluna CNJ)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Dados"
    ws.append(
        [
            "Número de CNJ",
            "Outro Número do Processo",
            "CLIENTE PRINCIPAL",
            "Responsável Principal",
            "STATUS",
        ]
    )
    # Bate pelo "Outro número" (pasta sem prefixo "Proc - ") -> CADASTRADO
    ws.append(
        [
            "",
            "0050900/001",
            "Cliente B",
            "Dr. Ciclano",
            "Novo",
        ]
    )
    # Bate pelo "Outro número" direto (processo fora do padrão CNJ) -> CADASTRADO
    ws.append(
        [
            "",
            "26.10.0846.001.00004-3",
            "Cliente C",
            "Dr. Ciclano",
            "Novo",
        ]
    )
    # Não bate -> PENDENTE
    ws.append(
        [
            "",
            "9999/2026",
            "Cliente Z",
            "Dr. Ciclano",
            "Novo",
        ]
    )
    wb.save(path)


def _criar_planilha_robo_protocolo(path: Path) -> None:
    """Planilha no formato exportado pelo "Robô Protocolo" (robô de
    e-mail/distribuição): aba única (sem aba CADASTRO), sem coluna de
    status, e com cabeçalhos próprios (Número do processo / Partes /
    Advogado responsável)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Planilha1"
    ws.append(
        [
            "Processado em",
            "Caixa monitorada",
            "Data/Hora do e-mail",
            "E-mail do destinatário",
            "Número do processo",
            "Partes",
            "Advogado responsável",
            "Tipo de distribuição",
            "Assunto original",
            "Link do e-mail",
        ]
    )
    # Bate com a base (mesmo CNJ) -> CADASTRADO
    ws.append(
        [
            "2026-10-09",
            "protocolocivel@deciofreire.com.br",
            "2026-10-08",
            "Protocolo Civel <protocolo@deciofreire.com.br>",
            "0010903-71.2018.5.03.0029",
            "Re: CONFERÊNCIA E DISTRIBUIÇÃO - AGRAVO DE INSTRUMENTO - Cliente A",
            "Dr. Fulano",
            "Agravo de Instrumento",
            "AGRAVO - Cliente A",
            "https://mail.google.com/mensagem-1",
        ]
    )
    # Não bate com nada -> PENDENTE
    ws.append(
        [
            "2026-10-10",
            "protocolocivel@deciofreire.com.br",
            "2026-10-09",
            "Protocolo Civel <protocolo@deciofreire.com.br>",
            "0766149-51.2026.8.18.0000",
            "Re: CONFERÊNCIA E DISTRIBUIÇÃO - AGRAVO DE INSTRUMENTO - Cliente X",
            "Dra. Sicrana",
            "Agravo de Instrumento",
            "AGRAVO - Cliente X",
            "https://mail.google.com/mensagem-2",
        ]
    )
    wb.save(path)


def _criar_export_ged(path: Path) -> None:
    df = pd.DataFrame(
        {
            "Anexado/gerado em": ["01/02/2026", "02/02/2026"],
            "Descrição/nome": [
                "planilha cadastro advogado",
                "planilha cadastro robo ti",
            ],
            "Tipo de arquivo/documento": [
                "Planilha de Cadastro - Advogado",
                "Planilha de Cadastro - Robô TI",
            ],
            "Contexto": ["Processo Judicial", "Processo Judicial"],
            "Pasta/Nome/ID": ["GED/123", "GED/124"],
        }
    )
    df.to_excel(path, index=False)


def _montar_fixtures(base_dir: Path) -> None:
    (base_dir / "entrada" / "base_legal_one").mkdir(parents=True, exist_ok=True)
    (base_dir / "entrada" / "planilhas_cadastro").mkdir(parents=True, exist_ok=True)
    (base_dir / "entrada" / "export_ged").mkdir(parents=True, exist_ok=True)

    _criar_base_legal_one(base_dir / "entrada" / "base_legal_one" / "base_legal_one.xlsx")
    _criar_planilha_cadastro_advogado(
        base_dir / "entrada" / "planilhas_cadastro" / "planilha cadastro advogado.xlsx"
    )
    _criar_planilha_cadastro_robo_ti(
        base_dir / "entrada" / "planilhas_cadastro" / "planilha cadastro robo ti.xlsx"
    )
    _criar_export_ged(base_dir / "entrada" / "export_ged" / "export_ged.xlsx")


def test_executar_batimento_classifica_status_corretamente(tmp_path: Path):
    _montar_fixtures(tmp_path)

    resultado = executar_batimento(tmp_path)

    assert resultado.planilhas_processadas == 2
    assert resultado.total_verificado == 6
    assert resultado.total_cadastrado == 3
    assert resultado.total_pendente == 2
    assert resultado.total_sem_numero == 1

    assert resultado.caminho_saida.exists()
    assert resultado.caminho_saida.parent == tmp_path / "saida"
    assert resultado.caminho_saida.name.startswith("Batimento_")

    df_detalhe = resultado.df_detalhe
    por_status = {
        (linha["Arquivo de origem"], linha["Número buscado"]): linha["Status"]
        for _, linha in df_detalhe.iterrows()
    }

    assert (
        por_status[("planilha cadastro advogado.xlsx", "0010903-71.2018.5.03.0029")]
        == STATUS_CADASTRADO
    )
    assert (
        por_status[("planilha cadastro advogado.xlsx", "0099999-99.2026.8.26.0100")]
        == STATUS_PENDENTE
    )
    assert por_status[("planilha cadastro advogado.xlsx", "")] == STATUS_SEM_NUMERO

    assert (
        por_status[("planilha cadastro robo ti.xlsx", "0050900/001")] == STATUS_CADASTRADO
    )
    assert (
        por_status[("planilha cadastro robo ti.xlsx", "26.10.0846.001.00004-3")]
        == STATUS_CADASTRADO
    )
    assert por_status[("planilha cadastro robo ti.xlsx", "9999/2026")] == STATUS_PENDENTE


def test_enriquecimento_com_export_ged(tmp_path: Path):
    _montar_fixtures(tmp_path)

    resultado = executar_batimento(tmp_path)
    df_detalhe = resultado.df_detalhe

    linha_advogado = df_detalhe[df_detalhe["Arquivo de origem"] == "planilha cadastro advogado.xlsx"].iloc[0]
    assert linha_advogado["Tipo"] == "Planilha de Cadastro - Advogado"
    assert linha_advogado["Data de envio"] == "01/02/2026"

    linha_robo_ti = df_detalhe[df_detalhe["Arquivo de origem"] == "planilha cadastro robo ti.xlsx"].iloc[0]
    assert linha_robo_ti["Tipo"] == "Planilha de Cadastro - Robô TI"
    assert linha_robo_ti["Data de envio"] == "02/02/2026"


def test_relatorio_gerado_tem_tres_abas_com_conteudo_esperado(tmp_path: Path):
    _montar_fixtures(tmp_path)

    resultado = executar_batimento(tmp_path)

    wb = load_workbook(resultado.caminho_saida)
    assert set(wb.sheetnames) == {"Resumo", "Pendencias", "Detalhe completo"}

    ws_pendencias = wb["Pendencias"]
    linhas_pendencias = list(ws_pendencias.iter_rows(values_only=True))
    cabecalho = linhas_pendencias[0]
    assert cabecalho == (
        "Arquivo de origem",
        "Tipo",
        "Responsável principal",
        "Cliente principal",
        "Número buscado",
        "Status",
        "Observação",
    )
    # 2 PENDENTE + 1 SEM NÚMERO = 3 linhas de dados na aba de pendências.
    assert len(linhas_pendencias) - 1 == 3

    ws_resumo = wb["Resumo"]
    linhas_resumo = list(ws_resumo.iter_rows(values_only=True))
    assert linhas_resumo[0][0] == "Arquivo de origem"
    assert len(linhas_resumo) - 1 == 2  # uma linha por arquivo de origem

    ws_detalhe = wb["Detalhe completo"]
    assert ws_detalhe.max_row - 1 == 6


def test_falha_amigavel_quando_base_legal_one_vazia(tmp_path: Path):
    (tmp_path / "entrada" / "base_legal_one").mkdir(parents=True, exist_ok=True)
    (tmp_path / "entrada" / "planilhas_cadastro").mkdir(parents=True, exist_ok=True)
    (tmp_path / "entrada" / "export_ged").mkdir(parents=True, exist_ok=True)

    try:
        executar_batimento(tmp_path)
        assert False, "Deveria ter levantado FileNotFoundError"
    except FileNotFoundError as erro:
        assert "base_legal_one" in str(erro)


def test_planilha_robo_protocolo_e_reconhecida(tmp_path: Path):
    """Planilha exportada pelo "Robô Protocolo" (aba única "Planilha1", sem
    aba CADASTRO, sem coluna de status, cabeçalhos "Número do processo" /
    "Partes" / "Advogado responsável") deve ser detectada e batida
    corretamente contra a base do Legal One."""
    (tmp_path / "entrada" / "base_legal_one").mkdir(parents=True, exist_ok=True)
    (tmp_path / "entrada" / "planilhas_cadastro").mkdir(parents=True, exist_ok=True)
    (tmp_path / "entrada" / "export_ged").mkdir(parents=True, exist_ok=True)

    _criar_base_legal_one(tmp_path / "entrada" / "base_legal_one" / "base_legal_one.xlsx")
    _criar_planilha_robo_protocolo(
        tmp_path / "entrada" / "planilhas_cadastro" / "planilha robo protocolo.xlsx"
    )

    resultado = executar_batimento(tmp_path)

    assert resultado.planilhas_processadas == 1
    assert resultado.total_verificado == 2
    assert resultado.total_cadastrado == 1
    assert resultado.total_pendente == 1
    assert resultado.total_sem_numero == 0

    df_detalhe = resultado.df_detalhe
    linhas = {
        linha["Número buscado"]: linha for _, linha in df_detalhe.iterrows()
    }

    linha_cadastrada = linhas["0010903-71.2018.5.03.0029"]
    assert linha_cadastrada["Status"] == STATUS_CADASTRADO
    assert linha_cadastrada["Cliente principal"] == (
        "Re: CONFERÊNCIA E DISTRIBUIÇÃO - AGRAVO DE INSTRUMENTO - Cliente A"
    )
    assert linha_cadastrada["Responsável principal"] == "Dr. Fulano"

    linha_pendente = linhas["0766149-51.2026.8.18.0000"]
    assert linha_pendente["Status"] == STATUS_PENDENTE
    assert linha_pendente["Cliente principal"] == (
        "Re: CONFERÊNCIA E DISTRIBUIÇÃO - AGRAVO DE INSTRUMENTO - Cliente X"
    )
    assert linha_pendente["Responsável principal"] == "Dra. Sicrana"

