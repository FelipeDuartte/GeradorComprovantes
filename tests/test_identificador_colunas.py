import pytest

from gerador_comprovantes.exceptions import ColunasNaoIdentificadasError
from gerador_comprovantes.services.identificador_colunas import IdentificadorDeColunas


def test_identifica_colunas_com_nomes_padrao():
    identificador = IdentificadorDeColunas()
    resultado = identificador.identificar(["Nome", "CPF", "Valor Total"])

    assert resultado.nome == "Nome"
    assert resultado.cpf == "CPF"
    assert resultado.valor == "Valor Total"


def test_prefere_valor_total_a_valor_parcela():
    identificador = IdentificadorDeColunas()
    resultado = identificador.identificar(
        ["Bolsista", "CPF", "Valor Parcela", "Valor Total Recebido"]
    )

    assert resultado.valor == "Valor Total Recebido"


def test_identifica_bolsista_como_nome():
    identificador = IdentificadorDeColunas()
    resultado = identificador.identificar(["Bolsista", "CPF", "Total Recebido"])

    assert resultado.nome == "Bolsista"


def test_levanta_erro_quando_nao_identifica_colunas():
    identificador = IdentificadorDeColunas()
    with pytest.raises(ColunasNaoIdentificadasError):
        identificador.identificar(["Coluna A", "Coluna B"])
