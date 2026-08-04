"""Configuração central de logging."""
from __future__ import annotations

import logging
from pathlib import Path

from gerador_comprovantes.config import Settings


def configurar_logging(base_dir: Path, settings: Settings) -> None:
    pasta_logs = base_dir / settings.pasta_logs
    pasta_logs.mkdir(exist_ok=True)

    logging.basicConfig(
        filename=pasta_logs / settings.nome_arquivo_log,
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        encoding="utf-8",
    )
