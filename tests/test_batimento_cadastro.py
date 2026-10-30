"""Testes do batimento entre planilhas de cadastro e a base do Legal One.

As fixtures são geradas programaticamente (nenhum dado real do usuário é
usado), montando a mesma estrutura de pastas esperada pelo robô
(``entrada/base_legal_one``, ``entrada/planilhas_cadastro``,
``entrada/export_ged``) dentro de um diretório temporário.
"""
import sys
from datetime import datetime
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


def _criar_base_legal_one_sem_titulo(path: Path) -> None:
    """Export novo do Legal One SEM linha de título solta: o cabeçalho real
    já está na linha 1 (ex: "Base Legal One até 30-10-2026.xlsx")."""
    wb = Workbook()
    ws = wb.active
    ws.append(
        [
            "Pasta",
            "Tipo processo/recurso/incidente",
            "Data do cadastro",
            "Número de CNJ",
            "Outro número",
            "Número antigo",
            "Escritório responsável/área",
            "Responsável principal (Coordenador)",
            "Cliente principal",
            "Grupo",
            "Tipo",
            "Status",
        ]
    )
    ws.append(
        [
            "Proc - 0070111/001",
            "Processo",
            "01/01/2026",
            "0010903-71.2018.5.03.0029",
            "",
            "",
            "Escritório X",
            "Dr. Fulano",
            "Cliente A",
            "Grupo 1",
            "Cível",
            "Ativo",
        ]
    )
    ws.append(
        [
            "Proc - 0070222/001",
            "Processo",
            "02/01/2026",
            "",
            "3344/2026",
            "",
            "Escritório X",
            "Dr. Fulano",
            "Cliente B",
            "Grupo 1",
            "Cível",
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
            # Valor datetime real (como vem do Excel), não string, para
            # cobrir a conversão em `carregar_planilha_cadastro`.
            datetime(2026, 10, 8),
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
            datetime(2026, 10, 9, 14, 30),
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


def _criar_planilha_simples_dois_colunas(path: Path) -> None:
    """Planilha simples de 2 colunas ("Processos:"/"Data:"), com aba de
    nome genérico/numérico (não "CADASTRO"), como os exports reais
    "Processos - Estratégico.xlsx"/"Processos - Tributário.xlsx"."""
    wb = Workbook()
    ws = wb.active
    ws.title = "90809908"
    ws.append(["Processos:", "Data:"])
    # Bate com a base (mesmo CNJ, só dígitos, com espaço em branco antes) ->
    # CADASTRADO.
    ws.append([" 00109037120185030029", datetime(2026, 1, 15)])
    # Não batem com nada -> PENDENTE.
    ws.append(["00000000120260000001", datetime(2026, 2, 20)])
    ws.append(["00000000220260000002", datetime(2026, 3, 5)])
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
        "Remetente do e-mail",
        "Data de envio do e-mail",
        "Data do processo",
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
    assert linha_cadastrada["Remetente do e-mail"] == "protocolocivel@deciofreire.com.br"
    assert linha_cadastrada["Data de envio do e-mail"] == "08/10/2026"
    # Formato Robô Protocolo não tem coluna de "data do processo" própria
    # (só data do e-mail); deve ficar vazia, sem erro.
    assert linha_cadastrada["Data do processo"] == ""

    linha_pendente = linhas["0766149-51.2026.8.18.0000"]
    assert linha_pendente["Status"] == STATUS_PENDENTE
    assert linha_pendente["Cliente principal"] == (
        "Re: CONFERÊNCIA E DISTRIBUIÇÃO - AGRAVO DE INSTRUMENTO - Cliente X"
    )
    assert linha_pendente["Responsável principal"] == "Dra. Sicrana"
    assert linha_pendente["Remetente do e-mail"] == "protocolocivel@deciofreire.com.br"
    # Hora diferente de meia-noite genuinamente presente na planilha de
    # origem deve ser preservada no formato "dd/mm/AAAA HH:MM".
    assert linha_pendente["Data de envio do e-mail"] == "09/10/2026 14:30"
    assert linha_pendente["Data do processo"] == ""


def test_planilha_simples_processos_data_e_reconhecida(tmp_path: Path):
    """Planilha simples de 2 colunas ("Processos:"/"Data:"), com aba de nome
    genérico/numérico (não "CADASTRO" nem nada parecido), como os exports
    reais "Processos - Estratégico.xlsx"/"Processos - Tributário.xlsx",
    deve ser detectada e batida corretamente contra a base do Legal One,
    extraindo também a "Data do processo" formatada."""
    (tmp_path / "entrada" / "base_legal_one").mkdir(parents=True, exist_ok=True)
    (tmp_path / "entrada" / "planilhas_cadastro").mkdir(parents=True, exist_ok=True)
    (tmp_path / "entrada" / "export_ged").mkdir(parents=True, exist_ok=True)

    _criar_base_legal_one(tmp_path / "entrada" / "base_legal_one" / "base_legal_one.xlsx")
    _criar_planilha_simples_dois_colunas(
        tmp_path / "entrada" / "planilhas_cadastro" / "processos estrategico.xlsx"
    )

    resultado = executar_batimento(tmp_path)

    assert resultado.planilhas_processadas == 1
    assert resultado.total_verificado == 3
    assert resultado.total_cadastrado == 1
    assert resultado.total_pendente == 2
    assert resultado.total_sem_numero == 0

    df_detalhe = resultado.df_detalhe
    linhas = {
        linha["Número buscado"]: linha for _, linha in df_detalhe.iterrows()
    }

    # Número com espaço em branco antes deve ser extraído/normalizado
    # corretamente (sem o espaço) e bater com a base do Legal One.
    linha_cadastrada = linhas["00109037120185030029"]
    assert linha_cadastrada["Status"] == STATUS_CADASTRADO
    assert linha_cadastrada["Data do processo"] == "15/01/2026"
    # Esse formato não tem coluna de remetente/data de e-mail: continua
    # vazio, sem erro.
    assert linha_cadastrada["Remetente do e-mail"] == ""
    assert linha_cadastrada["Data de envio do e-mail"] == ""

    linha_pendente_1 = linhas["00000000120260000001"]
    assert linha_pendente_1["Status"] == STATUS_PENDENTE
    assert linha_pendente_1["Data do processo"] == "20/02/2026"

    linha_pendente_2 = linhas["00000000220260000002"]
    assert linha_pendente_2["Status"] == STATUS_PENDENTE
    assert linha_pendente_2["Data do processo"] == "05/03/2026"


def test_colunas_remetente_e_data_email_ficam_vazias_sem_erro_no_formato_antigo(tmp_path: Path):
    """Planilhas no formato "Planilha de Cadastro - <tipo>" (advogado/robô TI)
    não têm nenhuma coluna de remetente/data de e-mail; as novas colunas
    devem simplesmente ficar vazias, sem erro."""
    _montar_fixtures(tmp_path)

    resultado = executar_batimento(tmp_path)

    df_detalhe = resultado.df_detalhe
    assert "Remetente do e-mail" in df_detalhe.columns
    assert "Data de envio do e-mail" in df_detalhe.columns
    assert "Data do processo" in df_detalhe.columns

    linhas_formato_antigo = df_detalhe[
        df_detalhe["Arquivo de origem"].isin(
            ["planilha cadastro advogado.xlsx", "planilha cadastro robo ti.xlsx"]
        )
    ]
    assert (linhas_formato_antigo["Remetente do e-mail"] == "").all()
    assert (linhas_formato_antigo["Data de envio do e-mail"] == "").all()
    assert (linhas_formato_antigo["Data do processo"] == "").all()


def test_base_legal_one_sem_linha_de_titulo_e_reconhecida(tmp_path: Path):
    """Export novo do Legal One cujo cabeçalho já está na linha 1 (sem a
    linha de título solto "usuario@empresa.com.br - ..." antes) deve ter o
    cabeçalho corretamente reconhecido, em vez de descartar a linha 1 cegamente
    com `skiprows=1` (o que faria a primeira linha de dados virar cabeçalho e
    nenhuma coluna-chave ser reconhecida, deixando a base "vazia" e todo
    processo aparecendo como PENDENTE)."""
    (tmp_path / "entrada" / "base_legal_one").mkdir(parents=True, exist_ok=True)
    (tmp_path / "entrada" / "planilhas_cadastro").mkdir(parents=True, exist_ok=True)
    (tmp_path / "entrada" / "export_ged").mkdir(parents=True, exist_ok=True)

    _criar_base_legal_one_sem_titulo(
        tmp_path / "entrada" / "base_legal_one" / "Base Legal One ate 30-10-2026.xlsx"
    )
    _criar_planilha_robo_protocolo(
        tmp_path / "entrada" / "planilhas_cadastro" / "planilha robo protocolo.xlsx"
    )

    resultado = executar_batimento(tmp_path)

    assert resultado.total_verificado == 2
    # As duas linhas da planilha de cadastro batem com a base (mesmo CNJ e
    # mesmo "Outro número" normalizado-para-dígitos), confirmando que o
    # cabeçalho da base foi corretamente reconhecido na linha 1.
    df_detalhe = resultado.df_detalhe
    linhas = {
        linha["Número buscado"]: linha for _, linha in df_detalhe.iterrows()
    }
    assert linhas["0010903-71.2018.5.03.0029"]["Status"] == STATUS_CADASTRADO


def test_base_legal_one_mista_com_e_sem_titulo_sao_reconhecidas_individualmente(tmp_path: Path):
    """Quando há múltiplos arquivos em `entrada/base_legal_one`, a detecção
    de linha de título solto deve ser feita por arquivo individualmente: um
    pode ter título solto e outro não."""
    (tmp_path / "entrada" / "base_legal_one").mkdir(parents=True, exist_ok=True)
    (tmp_path / "entrada" / "planilhas_cadastro").mkdir(parents=True, exist_ok=True)
    (tmp_path / "entrada" / "export_ged").mkdir(parents=True, exist_ok=True)

    _criar_base_legal_one(tmp_path / "entrada" / "base_legal_one" / "base_legal_one_antiga.xlsx")
    _criar_base_legal_one_sem_titulo(
        tmp_path / "entrada" / "base_legal_one" / "Base Legal One ate 30-10-2026.xlsx"
    )
    _criar_planilha_robo_protocolo(
        tmp_path / "entrada" / "planilhas_cadastro" / "planilha robo protocolo.xlsx"
    )

    resultado = executar_batimento(tmp_path)

    df_detalhe = resultado.df_detalhe
    linhas = {
        linha["Número buscado"]: linha for _, linha in df_detalhe.iterrows()
    }
    assert linhas["0010903-71.2018.5.03.0029"]["Status"] == STATUS_CADASTRADO
    assert linhas["0766149-51.2026.8.18.0000"]["Status"] == STATUS_PENDENTE


