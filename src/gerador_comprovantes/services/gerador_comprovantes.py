from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterator

from gerador_comprovantes.config import Settings
from gerador_comprovantes.exceptions import ModeloNaoEncontradoError
from gerador_comprovantes.models import Beneficiario
from gerador_comprovantes.services.conversor_pdf import ConversorPdf
from gerador_comprovantes.services.identificador_colunas import IdentificadorDeColunas
from gerador_comprovantes.services.leitor_excel import LeitorDeExcel
from gerador_comprovantes.services.preenchedor_word import PreenchedorWord
from gerador_comprovantes.services.resolvedor_caminhos import ResolvedorDeCaminhos

logger = logging.getLogger(__name__)

# Callback de progresso: (atual, total, mensagem) -> None
CallbackProgresso = Callable[[int, int, str], None]


@dataclass
class ResultadoGeracao:
    total_gerados: int
    pasta_saida: Path


class GeradorDeComprovantes:
    """
    Caso de uso principal: lê a planilha, preenche o modelo para cada
    beneficiário e converte cada um em PDF.

    Cada dependência (leitor de excel, identificador de colunas,
    preenchedor de word, conversor de pdf, resolvedor de caminhos) é
    injetada no construtor. Isso é Injeção de Dependência: facilita
    testes (basta injetar um "dublê"/mock de cada serviço) e permite
    trocar implementações sem alterar esta classe.
    """

    def __init__(
        self,
        settings: Settings,
        resolvedor_caminhos: ResolvedorDeCaminhos,
        leitor_excel: LeitorDeExcel,
        identificador_colunas: IdentificadorDeColunas,
        preenchedor_word: PreenchedorWord,
        conversor_pdf: ConversorPdf,
    ) -> None:
        self._settings = settings
        self._caminhos = resolvedor_caminhos
        self._leitor_excel = leitor_excel
        self._identificador_colunas = identificador_colunas
        self._preenchedor_word = preenchedor_word
        self._conversor_pdf = conversor_pdf

    def gerar(self, on_progresso: CallbackProgresso | None = None) -> ResultadoGeracao:
        notificar = on_progresso or (lambda atual, total, msg: None)

        notificar(0, 0, "Verificando LibreOffice...")
        if not self._conversor_pdf.disponivel():
            from gerador_comprovantes.exceptions import LibreOfficeNaoEncontradoError

            raise LibreOfficeNaoEncontradoError(
                "LibreOffice não encontrado. Instale-o e tente novamente."
            )

        caminho_dados = self._caminhos.caminho_absoluto(
            self._settings.pasta_dados, self._settings.nome_arquivo_dados
        )
        caminho_modelo = self._localizar_modelo()

        logger.info("Base: %s", self._caminhos.base_dir)
        logger.info("Planilha: %s", caminho_dados)
        logger.info("Modelo: %s", caminho_modelo)

        df = self._leitor_excel.ler(caminho_dados)
        colunas = self._identificador_colunas.identificar(list(df.columns))

        pasta_saida = self._caminhos.base_dir / self._settings.pasta_saida
        pasta_temp = pasta_saida / self._settings.pasta_temp
        pasta_saida.mkdir(exist_ok=True)
        pasta_temp.mkdir(exist_ok=True)

        total = len(df)
        notificar(0, total, "Iniciando...")

        for indice, beneficiario in enumerate(self._iterar_beneficiarios(df, colunas), start=1):
            self._gerar_um_comprovante(beneficiario, caminho_modelo, pasta_temp, pasta_saida)
            notificar(indice, total, f"Gerado {indice} de {total}...")

        pasta_temp.rmdir()
        logger.info("Geração concluída: %s comprovantes em %s", total, pasta_saida)

        return ResultadoGeracao(total_gerados=total, pasta_saida=pasta_saida)

    def _iterar_beneficiarios(self, df, colunas) -> Iterator[Beneficiario]:
        for _, linha in df.iterrows():
            yield Beneficiario.a_partir_de_linha(
                nome=linha[colunas.nome],
                cpf=linha[colunas.cpf],
                valor=linha[colunas.valor],
            )

    def _gerar_um_comprovante(
        self,
        beneficiario: Beneficiario,
        caminho_modelo: Path,
        pasta_temp: Path,
        pasta_saida: Path,
    ) -> None:
        docx_temp = pasta_temp / f"{beneficiario.nome_arquivo_seguro()}.docx"

        self._preenchedor_word.preencher_e_salvar(
            caminho_modelo,
            beneficiario.para_mapa_substituicao(self._settings.placeholders),
            docx_temp,
        )
        try:
            self._conversor_pdf.converter(docx_temp, pasta_saida)
        finally:
            docx_temp.unlink(missing_ok=True)

    def _localizar_modelo(self) -> Path:
        for nome in self._settings.nomes_possiveis_modelo:
            caminho = self._caminhos.caminho_absoluto(self._settings.pasta_modelos, nome)
            if caminho.exists():
                return caminho

        raise ModeloNaoEncontradoError(
            "Modelo não encontrado. Verifique se há um arquivo .docx de modelo "
            f"na pasta '{self._settings.pasta_modelos}'."
        )
