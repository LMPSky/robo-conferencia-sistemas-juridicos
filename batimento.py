#!/usr/bin/env python
"""Ponto de entrada do robô de batimento de cadastro de processos.

Uso:
    python batimento.py

Lê as planilhas em `entrada/base_legal_one`, `entrada/planilhas_cadastro` e
(opcionalmente) `entrada/export_ged`, e grava o relatório de batimento em
`saida/Batimento_AAAAMMDD_HHMMSS.xlsx`.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from robo_conferencia.batimento_cadastro import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
