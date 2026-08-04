"""
Identificação automática de colunas na planilha.

Foi transformado em classe porque encapsula regras de pontuação que
podem crescer (novos sinônimos, novos pesos) e porque isso permite,
por exemplo, criar variações do identificador (outro idioma, outras
palavras-chave) sem tocar no resto do sistema — basta injetar outra
instância.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from gerador_comprovantes.exceptions import ColunasNaoIdentificadasError
from gerador_comprovantes.services.formatadores import normalizar


@dataclass
class ColunasIdentificadas:
    """Resultado da identificação: nome das colunas reais na planilha."""

    nome: str
    cpf: str
    valor: str


@dataclass
class IdentificadorDeColunas:
    """
    Identifica, por heurística, quais colunas de um DataFrame correspondem
    a nome, CPF e valor — mesmo que a planilha não siga um padrão fixo
    de nomenclatura.
    """

    palavras_chave_nome: tuple[str, ...] = ("nome", "bolsista")
    palavras_chave_cpf: tuple[str, ...] = ("cpf",)

    # peso de cada termo ao pontuar a coluna de valor
    pesos_valor: dict[str, int] = field(
        default_factory=lambda: {
            "valor total": 5,
            "total recebido": 5,
            "recebido": 3,
            "total": 2,
            "valor": 1,
        }
    )
    penalidades_valor: tuple[str, ...] = ("parcela", "mensal")
    peso_penalidade: int = -3

    def identificar(self, colunas: list[str]) -> ColunasIdentificadas:
        coluna_nome = self._buscar_por_palavras_chave(colunas, self.palavras_chave_nome)
        coluna_cpf = self._buscar_por_palavras_chave(colunas, self.palavras_chave_cpf)
        coluna_valor = self._buscar_melhor_coluna_valor(colunas)

        if not all([coluna_nome, coluna_cpf, coluna_valor]):
            raise ColunasNaoIdentificadasError(
                "Não foi possível identificar automaticamente as colunas.\n\n"
                "Verifique se a planilha possui colunas de Nome, CPF e Valor."
            )

        return ColunasIdentificadas(nome=coluna_nome, cpf=coluna_cpf, valor=coluna_valor)

    def _buscar_por_palavras_chave(
        self, colunas: list[str], palavras_chave: tuple[str, ...]
    ) -> str | None:
        for coluna in colunas:
            coluna_normalizada = normalizar(coluna)
            if any(palavra in coluna_normalizada for palavra in palavras_chave):
                return coluna
        return None

    def _buscar_melhor_coluna_valor(self, colunas: list[str]) -> str | None:
        candidatas: list[tuple[int, str]] = []

        for coluna in colunas:
            coluna_normalizada = normalizar(coluna)
            if not any(
                termo in coluna_normalizada
                for termo in ("valor", "total", "recebido")
            ):
                continue

            pontuacao = sum(
                peso
                for termo, peso in self.pesos_valor.items()
                if termo in coluna_normalizada
            )
            if any(termo in coluna_normalizada for termo in self.penalidades_valor):
                pontuacao += self.peso_penalidade

            candidatas.append((pontuacao, coluna))

        if not candidatas:
            return None

        candidatas.sort(reverse=True)
        return candidatas[0][1]
