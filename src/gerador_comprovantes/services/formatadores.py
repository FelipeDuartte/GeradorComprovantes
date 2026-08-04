from __future__ import annotations

import re
import unicodedata


def normalizar(texto: str) -> str:
    """Remove acentos e caixa alta para permitir comparação de texto tolerante."""
    texto = str(texto).lower()
    texto = unicodedata.normalize("NFD", texto)
    return "".join(c for c in texto if unicodedata.category(c) != "Mn")


def formatar_cpf(cpf: str | float | int) -> str:
    """Formata um CPF para o padrão 000.000.000-00 quando possível."""
    digitos = re.sub(r"\D", "", str(cpf))
    if len(digitos) == 11:
        return f"{digitos[:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:]}"
    return digitos


def formatar_valor(valor: float) -> str:
    """Formata um número no padrão monetário brasileiro: R$ 1.234,56."""
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
