"""
Camada de apresentação (GUI).
"""
from __future__ import annotations

import threading
import tkinter as tk
from tkinter import messagebox, ttk

from gerador_comprovantes.exceptions import GeradorComprovantesError
from gerador_comprovantes.services.gerador_comprovantes import GeradorDeComprovantes


class JanelaPrincipal:
    def __init__(self, gerador: GeradorDeComprovantes) -> None:
        self._gerador = gerador
        self._janela = tk.Tk()
        self._montar_widgets()

    def executar(self) -> None:
        self._janela.mainloop()

    def _montar_widgets(self) -> None:
        self._janela.title("Gerador de Comprovantes")
        self._janela.geometry("420x280")
        self._janela.resizable(False, False)

        tk.Label(
            self._janela,
            text="Gerador de Comprovantes",
            font=("Segoe UI", 15, "bold"),
        ).pack(pady=20)

        self._barra = ttk.Progressbar(
            self._janela, orient="horizontal", length=300, mode="determinate"
        )
        self._barra.pack(pady=15)

        self._status_label = tk.Label(self._janela, text="Aguardando...", font=("Segoe UI", 10))
        self._status_label.pack()

        tk.Label(
            self._janela,
            text="Os comprovantes serão gerados em PDF",
            font=("Segoe UI", 9),
            fg="gray",
        ).pack(pady=(4, 0))

        self._botao = tk.Button(
            self._janela,
            text="Gerar comprovantes",
            font=("Segoe UI", 11),
            width=28,
            height=2,
            command=self._iniciar_geracao,
        )
        self._botao.pack(pady=16)

    def _iniciar_geracao(self) -> None:
        self._botao.config(state="disabled")
        self._barra["value"] = 0
        self._status_label.config(text="Iniciando...")

        thread = threading.Thread(target=self._gerar_em_background, daemon=True)
        thread.start()

    def _gerar_em_background(self) -> None:
        try:
            resultado = self._gerador.gerar(on_progresso=self._on_progresso)
            self._janela.after(0, self._exibir_sucesso, resultado.total_gerados, str(resultado.pasta_saida))
        except GeradorComprovantesError as erro:
            self._janela.after(0, self._exibir_erro, str(erro))
        except Exception as erro:  # erro inesperado, ainda assim não deve travar a UI
            self._janela.after(0, self._exibir_erro, f"Erro inesperado: {erro}")

    def _on_progresso(self, atual: int, total: int, mensagem: str) -> None:
        def atualizar() -> None:
            if total:
                self._barra["maximum"] = total
                self._barra["value"] = atual
            self._status_label.config(text=mensagem)

        self._janela.after(0, atualizar)

    def _exibir_sucesso(self, total: int, pasta_saida: str) -> None:
        self._status_label.config(text="Concluído ✔")
        self._botao.config(state="normal")
        messagebox.showinfo(
            "Sucesso", f"{total} comprovante(s) gerado(s) com sucesso!\n\nConfira a pasta:\n{pasta_saida}"
        )

    def _exibir_erro(self, mensagem: str) -> None:
        self._status_label.config(text="Erro ao gerar.")
        self._botao.config(state="normal")
        messagebox.showerror("Erro", mensagem)
