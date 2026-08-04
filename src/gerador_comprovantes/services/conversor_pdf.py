"""
Conversão de .docx para .pdf.
"""
from __future__ import annotations

import platform
import shutil
import subprocess
from abc import ABC, abstractmethod
from pathlib import Path

from gerador_comprovantes.exceptions import ConversaoPdfError, LibreOfficeNaoEncontradoError


class ConversorPdf(ABC):
    """Contrato que qualquer conversor de docx -> pdf deve seguir."""

    @abstractmethod
    def disponivel(self) -> bool:
        """Indica se o conversor está pronto para uso no sistema atual."""

    @abstractmethod
    def converter(self, caminho_docx: Path, pasta_saida: Path) -> Path:
        """Converte o arquivo e retorna o caminho do PDF gerado."""


class ConversorLibreOffice(ConversorPdf):
    """
    Converte documentos usando o LibreOffice em modo headless.

    A busca pelo executável é multiplataforma:
    1. Primeiro tenta `shutil.which`, que respeita o PATH do sistema
       operacional (funciona bem em Linux e em instalações customizadas).
    2. Se não encontrar, cai para os caminhos de instalação padrão de
       cada sistema operacional.
    """

    CAMINHOS_PADRAO: dict[str, tuple[str, ...]] = {
        "Darwin": (
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/usr/local/bin/soffice",
        ),
        "Linux": (
            "/usr/bin/soffice",
            "/usr/local/bin/soffice",
            "/snap/bin/libreoffice",
        ),
        "Windows": (
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ),
    }

    NOMES_EXECUTAVEL: dict[str, str] = {
        "Windows": "soffice.exe",
    }

    def __init__(self, timeout_segundos: int = 60) -> None:
        self._timeout = timeout_segundos
        self._caminho_executavel: Path | None = None

    def disponivel(self) -> bool:
        return self._localizar_executavel() is not None

    def converter(self, caminho_docx: Path, pasta_saida: Path) -> Path:
        soffice = self._localizar_executavel()
        if soffice is None:
            raise LibreOfficeNaoEncontradoError(
                "LibreOffice não encontrado neste computador.\n\n"
                "Instale o LibreOffice (gratuito) e tente novamente:\n"
                "https://www.libreoffice.org/download/download/"
            )

        try:
            resultado = subprocess.run(
                [
                    str(soffice),
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(pasta_saida),
                    str(caminho_docx),
                ],
                capture_output=True,
                text=True,
                timeout=self._timeout,
            )
        except subprocess.TimeoutExpired as erro:
            raise ConversaoPdfError(
                f"Tempo limite excedido ao converter '{caminho_docx.name}' para PDF."
            ) from erro
        except OSError as erro:
            raise ConversaoPdfError(
                f"Erro do sistema ao executar o LibreOffice: {erro}"
            ) from erro

        if resultado.returncode != 0:
            raise ConversaoPdfError(
                f"Erro ao converter '{caminho_docx.name}' para PDF:\n{resultado.stderr}"
            )

        return pasta_saida / f"{caminho_docx.stem}.pdf"

    def _localizar_executavel(self) -> Path | None:
        if self._caminho_executavel is not None:
            return self._caminho_executavel

        sistema = platform.system()  # "Darwin", "Linux" ou "Windows"
        nome_no_path = self.NOMES_EXECUTAVEL.get(sistema, "soffice")

        encontrado_no_path = shutil.which(nome_no_path)
        if encontrado_no_path:
            self._caminho_executavel = Path(encontrado_no_path)
            return self._caminho_executavel

        for caminho_str in self.CAMINHOS_PADRAO.get(sistema, ()):
            caminho = Path(caminho_str)
            if caminho.exists():
                self._caminho_executavel = caminho
                return self._caminho_executavel

        return None
