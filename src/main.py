from gerador_comprovantes.config import SETTINGS
from gerador_comprovantes.gui.app import JanelaPrincipal
from gerador_comprovantes.logging_config import configurar_logging
from gerador_comprovantes.services.conversor_pdf import ConversorLibreOffice
from gerador_comprovantes.services.gerador_comprovantes import GeradorDeComprovantes
from gerador_comprovantes.services.identificador_colunas import IdentificadorDeColunas
from gerador_comprovantes.services.leitor_excel import LeitorDeExcel
from gerador_comprovantes.services.preenchedor_word import PreenchedorWord
from gerador_comprovantes.services.resolvedor_caminhos import ResolvedorDeCaminhos


def main() -> None:
    resolvedor_caminhos = ResolvedorDeCaminhos()
    configurar_logging(resolvedor_caminhos.base_dir, SETTINGS)

    gerador = GeradorDeComprovantes(
        settings=SETTINGS,
        resolvedor_caminhos=resolvedor_caminhos,
        leitor_excel=LeitorDeExcel(),
        identificador_colunas=IdentificadorDeColunas(),
        preenchedor_word=PreenchedorWord(),
        conversor_pdf=ConversorLibreOffice(timeout_segundos=SETTINGS.timeout_conversao_segundos),
    )

    JanelaPrincipal(gerador).executar()


if __name__ == "__main__":
    main()
