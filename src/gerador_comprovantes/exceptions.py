"""
Exceções específicas do domínio.

Ter exceções próprias (em vez de usar Exception genérica) permite que
cada camada trate apenas os erros que sabe tratar, e que a UI monte
mensagens amigáveis sem precisar interpretar strings de erro.
"""


class GeradorComprovantesError(Exception):
    """Classe base para todos os erros conhecidos da aplicação."""


class LibreOfficeNaoEncontradoError(GeradorComprovantesError):
    """LibreOffice não foi localizado no sistema operacional atual."""


class ConversaoPdfError(GeradorComprovantesError):
    """Falha ao converter um .docx para .pdf."""


class ArquivoDadosNaoEncontradoError(GeradorComprovantesError):
    """A planilha de dados (Excel) não foi encontrada."""


class ModeloNaoEncontradoError(GeradorComprovantesError):
    """O modelo .docx do comprovante não foi encontrado."""


class ColunasNaoIdentificadasError(GeradorComprovantesError):
    """Não foi possível identificar automaticamente as colunas necessárias."""
