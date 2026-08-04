"""
Teste do orquestrador principal usando dublês (mocks) no lugar das
dependências reais (Excel, Word, PDF). 
"""
from pathlib import Path
from unittest.mock import MagicMock

import pandas as pd
import pytest

from gerador_comprovantes.config import Settings
from gerador_comprovantes.exceptions import LibreOfficeNaoEncontradoError, ModeloNaoEncontradoError
from gerador_comprovantes.services.gerador_comprovantes import GeradorDeComprovantes
from gerador_comprovantes.services.identificador_colunas import IdentificadorDeColunas


@pytest.fixture
def settings() -> Settings:
    return Settings()


@pytest.fixture
def dependencias_mockadas(tmp_path, settings):
    resolvedor = MagicMock()
    resolvedor.base_dir = tmp_path
    resolvedor.caminho_absoluto.side_effect = lambda *partes: tmp_path.joinpath(*partes)

    (tmp_path / settings.pasta_modelos).mkdir()
    caminho_modelo = tmp_path / settings.pasta_modelos / settings.nomes_possiveis_modelo[0]
    caminho_modelo.touch()

    leitor_excel = MagicMock()
    leitor_excel.ler.return_value = pd.DataFrame(
        {"Nome": ["Ana"], "CPF": ["12345678901"], "Valor Total": [100.0]}
    )

    preenchedor_word = MagicMock()
    conversor_pdf = MagicMock()
    conversor_pdf.disponivel.return_value = True

    return {
        "resolvedor": resolvedor,
        "leitor_excel": leitor_excel,
        "preenchedor_word": preenchedor_word,
        "conversor_pdf": conversor_pdf,
    }


def _montar_gerador(settings, deps) -> GeradorDeComprovantes:
    return GeradorDeComprovantes(
        settings=settings,
        resolvedor_caminhos=deps["resolvedor"],
        leitor_excel=deps["leitor_excel"],
        identificador_colunas=IdentificadorDeColunas(),
        preenchedor_word=deps["preenchedor_word"],
        conversor_pdf=deps["conversor_pdf"],
    )


def test_gera_um_comprovante_por_linha(settings, dependencias_mockadas):
    gerador = _montar_gerador(settings, dependencias_mockadas)

    resultado = gerador.gerar()

    assert resultado.total_gerados == 1
    dependencias_mockadas["preenchedor_word"].preencher_e_salvar.assert_called_once()
    dependencias_mockadas["conversor_pdf"].converter.assert_called_once()


def test_notifica_progresso_via_callback(settings, dependencias_mockadas):
    gerador = _montar_gerador(settings, dependencias_mockadas)
    chamadas = []

    gerador.gerar(on_progresso=lambda atual, total, msg: chamadas.append((atual, total, msg)))

    assert (1, 1, "Gerado 1 de 1...") in chamadas


def test_levanta_erro_quando_libreoffice_indisponivel(settings, dependencias_mockadas):
    dependencias_mockadas["conversor_pdf"].disponivel.return_value = False
    gerador = _montar_gerador(settings, dependencias_mockadas)

    with pytest.raises(LibreOfficeNaoEncontradoError):
        gerador.gerar()


def test_levanta_erro_quando_modelo_nao_existe(settings, dependencias_mockadas, tmp_path):
    (tmp_path / settings.pasta_modelos / settings.nomes_possiveis_modelo[0]).unlink()
    gerador = _montar_gerador(settings, dependencias_mockadas)

    with pytest.raises(ModeloNaoEncontradoError):
        gerador.gerar()
