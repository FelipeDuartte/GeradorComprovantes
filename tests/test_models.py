from gerador_comprovantes.models import Beneficiario


def test_beneficiario_formata_cpf_e_valor_automaticamente():
    b = Beneficiario.a_partir_de_linha(nome="Maria Silva", cpf="12345678901", valor=1500.5)

    assert b.nome == "Maria Silva"
    assert b.cpf_formatado == "123.456.789-01"
    assert b.valor_formatado == "R$ 1.500,50"


def test_nome_arquivo_seguro_remove_caracteres_invalidos():
    b = Beneficiario.a_partir_de_linha(nome='João "Silva"/Costa', cpf="12345678901", valor=10)

    assert b.nome_arquivo_seguro() == 'João _Silva__Costa'


def test_para_mapa_substituicao():
    b = Beneficiario.a_partir_de_linha(nome="Ana", cpf="12345678901", valor=10)
    placeholders = {"nome": "{{Nome}}", "cpf": "{{CPF}}", "valor": "{{Valor}}"}

    mapa = b.para_mapa_substituicao(placeholders)

    assert mapa == {
        "{{Nome}}": "Ana",
        "{{CPF}}": "123.456.789-01",
        "{{Valor}}": "R$ 10,00",
    }
