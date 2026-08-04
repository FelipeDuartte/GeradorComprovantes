"""
Modelos de domínio.

`Beneficiario` representa a entidade central do sistema: a pessoa que
vai receber o comprovante. Encapsular isso em uma classe (em vez de
passar `nome`, `cpf`, `valor` soltos entre funções) deixa as
assinaturas mais limpas e centraliza regras de formatação/validação.
"""
from __future__ import annotations

from dataclasses import dataclass

from gerador_comprovantes.services.formatadores import formatar_cpf, formatar_valor


@dataclass(frozen=True)
class Beneficiario:
    """Representa uma pessoa que receberá um comprovante."""

    nome: str
    cpf_formatado: str
    valor_formatado: str

    @classmethod
    def a_partir_de_linha(cls, nome: str, cpf: str, valor: float) -> "Beneficiario":
        """
        Cria um Beneficiario já formatando CPF e valor.

        Manter a formatação aqui (e não espalhada no orquestrador)
        garante que sempre que um Beneficiario existir, seus dados
        já estarão prontos para exibição.
        """
        return cls(
            nome=str(nome).strip(),
            cpf_formatado=formatar_cpf(cpf),
            valor_formatado=formatar_valor(valor),
        )

    def nome_arquivo_seguro(self) -> str:
        """Retorna o nome sem caracteres inválidos para nomes de arquivo."""
        import re

        return re.sub(r'[\\/*?:"<>|]', "_", self.nome)

    def para_mapa_substituicao(self, placeholders: dict[str, str]) -> dict[str, str]:
        """Monta o dicionário {placeholder: valor} usado no preenchimento do docx."""
        return {
            placeholders["nome"]: self.nome,
            placeholders["cpf"]: self.cpf_formatado,
            placeholders["valor"]: self.valor_formatado,
        }
