"""
Resolução de caminhos, funcionando em Windows, macOS e Linux, tanto
rodando como script quanto empacotado (PyInstaller).
"""
from __future__ import annotations

import sys
from pathlib import Path


class ResolvedorDeCaminhos:
    """Encontra a pasta-base da aplicação, independente de SO ou empacotamento."""

    def __init__(self) -> None:
        self._base_dir = self._descobrir_base_dir()

    @property
    def base_dir(self) -> Path:
        return self._base_dir

    def caminho_absoluto(self, *partes: str) -> Path:
        """Monta um caminho absoluto a partir da pasta-base da aplicação."""
        return self._base_dir.joinpath(*partes)

    @staticmethod
    def _descobrir_base_dir() -> Path:
        rodando_empacotado = getattr(sys, "frozen", False)

        if not rodando_empacotado:
            # Rodando como script: sobe 2 níveis a partir deste arquivo
            # (services/ -> gerador_comprovantes/ -> src/ -> raiz do projeto)
            return Path(__file__).resolve().parents[3]

        executavel = Path(sys.executable).resolve()

        if ".app/Contents/MacOS" in str(executavel):
            # Bundle .app do macOS: a pasta-base é a que CONTÉM o .app
            app_path = Path(str(executavel).split(".app/Contents/MacOS")[0] + ".app")
            return app_path.parent

        # Windows e Linux empacotados: o executável fica direto na pasta-base
        return executavel.parent
