"""
Configurações centralizadas.

Manter constantes e caminhos em um único lugar evita "magic strings"
espalhadas pelo código e facilita ajustar comportamento sem caçar
valores em vários arquivos.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    """
    Configurações de nomes de pastas/arquivos usados pela aplicação.

    `frozen=True` torna a instância imutável: uma vez criada, não pode
    ser alterada acidentalmente em outro ponto do código.
    """

    pasta_dados: str = "dados"
    nome_arquivo_dados: str = "dadosteste.xlsx"

    pasta_modelos: str = "modelos"
    nomes_possiveis_modelo: tuple[str, ...] = (
        "MODELO_COMPROVANTE.docx",
        "MODELO COMPROVANTE.docx",
        "MODELO COMPROVANTE RENDIMENTOS.docx",
    )

    pasta_saida: str = "output"
    pasta_temp: str = "_temp_docx"

    pasta_logs: str = "logs"
    nome_arquivo_log: str = "automacao.log"

    timeout_conversao_segundos: int = 60

    placeholders: dict = field(
        default_factory=lambda: {
            "nome": "{{Nome}}",
            "cpf": "{{CPF}}",
            "valor": "{{Valor}}",
        }
    )


SETTINGS = Settings()
