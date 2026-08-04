"""Leitura da planilha de entrada."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from gerador_comprovantes.exceptions import ArquivoDadosNaoEncontradoError


class LeitorDeExcel:
    """Responsável apenas por ler e normalizar a planilha de entrada."""

    def ler(self, caminho: Path) -> pd.DataFrame:
        if not caminho.exists():
            raise ArquivoDadosNaoEncontradoError(
                f"Arquivo de dados não encontrado:\n{caminho}"
            )

        df = pd.read_excel(caminho, engine="openpyxl")
        df.columns = df.columns.str.strip()
        return df
