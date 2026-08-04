from gerador_comprovantes.services.formatadores import formatar_cpf, formatar_valor, normalizar


def test_formatar_cpf_com_11_digitos():
    assert formatar_cpf("12345678901") == "123.456.789-01"


def test_formatar_cpf_ja_formatado():
    assert formatar_cpf("123.456.789-01") == "123.456.789-01"


def test_formatar_cpf_com_tamanho_invalido_retorna_so_digitos():
    assert formatar_cpf("123") == "123"


def test_formatar_valor_monetario_brasileiro():
    assert formatar_valor(1234.5) == "R$ 1.234,50"


def test_formatar_valor_sem_milhar():
    assert formatar_valor(50) == "R$ 50,00"


def test_normalizar_remove_acentos_e_caixa():
    assert normalizar("VALOR TOTAL RECEBIDO") == "valor total recebido"
    assert normalizar("Não Íntegro") == "nao integro"
