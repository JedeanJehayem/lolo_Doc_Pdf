import os
import customtkinter as ctk
import sys
import subprocess
import threading
from tkinter import filedialog, messagebox

from backend import (
    criar_estrutura_pastas_modelo,
    localizar_arquivos_modelo,
    localizar_arquivos_conversao,
    executar_processamento_modelo,
    executar_conversao_word_para_pdf,
    contar_campos_word,
    contar_colunas_excel,
)


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Gerador de Documentos")

        self.LARGURA_JANELA = 980
        self.ALTURA_JANELA = 680

        self.geometry(f"{self.LARGURA_JANELA}x{self.ALTURA_JANELA}")
        self.minsize(900, 620)

        self.pasta_var = ctk.StringVar(value="")
        self.quantidade_campos_var = ctk.IntVar(value=3)
        self.gerar_pdf_var = ctk.BooleanVar(value=True)
        self.modo_var = ctk.StringVar(value="modelo")

        self.after(50, self.centralizar_janela)

        self._montar_layout()
        self.atualizar_modo()

    def centralizar_janela(self):
        self.update_idletasks()

        largura = self.LARGURA_JANELA
        altura = self.ALTURA_JANELA

        largura_tela = self.winfo_screenwidth()
        altura_tela = self.winfo_screenheight()

        pos_x = max((largura_tela // 2) - (largura // 2), 0)
        pos_y = max((altura_tela // 2) - (altura // 2), 0)

        self.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    def normalizar_pasta_raiz(self, pasta):
        pasta = pasta.strip()
        if not pasta:
            return ""

        pasta = os.path.normpath(pasta)

        nome_final = os.path.basename(pasta).lower()
        if nome_final in {"entrada", "base", "saida", "word", "pdf"}:
            pasta = os.path.dirname(pasta)

        return pasta

    def on_slider_change(self, value):
        qtd = int(round(value))
        self.quantidade_campos_var.set(qtd)
        self.lbl_qtd.configure(text=f"{qtd} parâmetro{'s' if qtd > 1 else ''}")

    def escolher_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta principal")
        if pasta:
            pasta = self.normalizar_pasta_raiz(pasta)
            self.pasta_var.set(pasta)

    def criar_pastas(self):
        try:
            pasta_raiz = self.normalizar_pasta_raiz(self.pasta_var.get())

            if not pasta_raiz:
                messagebox.showwarning(
                    "Atenção",
                    "Selecione uma pasta antes de criar a estrutura."
                )
                return

            self.pasta_var.set(pasta_raiz)

            resultado = criar_estrutura_pastas_modelo(pasta_raiz)

            if resultado["criou_alguma"]:
                messagebox.showinfo("Sucesso", "Estrutura de pastas criada com sucesso.")
            else:
                messagebox.showinfo("Informação", "Estrutura já existente.")

        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def atualizar_modo(self):
        modo = self.modo_var.get()

        self.entry_arquivo_1.configure(state="normal")
        self.entry_arquivo_1.delete(0, "end")
        self.entry_arquivo_1.configure(state="readonly")

        self.entry_arquivo_2.configure(state="normal")
        self.entry_arquivo_2.delete(0, "end")
        self.entry_arquivo_2.configure(state="readonly")

        if modo == "modelo":
            self.frame_parametros.grid()
            self.chk_pdf.grid()

            self.lbl_arquivo_1.grid()
            self.entry_arquivo_1.grid()
            self.lbl_arquivo_2.grid()
            self.entry_arquivo_2.grid()

            self.lbl_arquivo_1.configure(text="Primeiro Excel encontrado")
            self.lbl_arquivo_2.configure(text="Primeiro Word da base encontrado")

        else:
            self.frame_parametros.grid_remove()
            self.chk_pdf.grid_remove()

            self.lbl_arquivo_1.grid()
            self.entry_arquivo_1.grid()
            self.lbl_arquivo_2.grid_remove()
            self.entry_arquivo_2.grid_remove()

            self.lbl_arquivo_1.configure(text="Primeiro Word encontrado")

    def localizar_arquivos_ui(self):
        try:
            modo = self.modo_var.get()
            pasta = self.normalizar_pasta_raiz(self.pasta_var.get().strip())

            if not pasta:
                messagebox.showwarning(
                    "Atenção",
                    "Selecione uma pasta antes de localizar os arquivos."
                )
                return

            self.pasta_var.set(pasta)

            if modo == "modelo":
                resultado = localizar_arquivos_modelo(pasta)

                self.entry_arquivo_1.configure(state="normal")
                self.entry_arquivo_1.delete(0, "end")
                self.entry_arquivo_1.insert(0, resultado["arquivo_excel"])
                self.entry_arquivo_1.configure(state="readonly")

                self.entry_arquivo_2.configure(state="normal")
                self.entry_arquivo_2.delete(0, "end")
                self.entry_arquivo_2.insert(0, resultado["arquivo_word"])
                self.entry_arquivo_2.configure(state="readonly")

                info_word = contar_campos_word(resultado["arquivo_word"])
                info_excel = contar_colunas_excel(resultado["arquivo_excel"])

                total_word = info_word["total_campos"]
                total_excel = info_excel["total_colunas"]
                total_param = self.quantidade_campos_var.get()

                mensagens = []

                if total_param > total_word:
                    mensagens.append(
                        f"• Você selecionou {total_param} parâmetro(s), mas o Word tem apenas {total_word} campo(s)."
                    )

                if total_param > total_excel:
                    mensagens.append(
                        f"• Você selecionou {total_param} parâmetro(s), mas o Excel tem apenas {total_excel} coluna(s)."
                    )

                if total_param < total_word:
                    mensagens.append(
                        f"• O Word tem {total_word} campo(s), mas você selecionou apenas {total_param} parâmetro(s)."
                    )

                if total_param < total_excel:
                    mensagens.append(
                        f"• O Excel tem {total_excel} coluna(s), mas você selecionou apenas {total_param} parâmetro(s)."
                    )

                if mensagens:
                    messagebox.showwarning(
                        "Validação de parâmetros",
                        "Atenção:\n\n" + "\n".join(mensagens)
                    )

                messagebox.showinfo(
                    "Arquivos encontrados",
                    "Primeiro Excel e primeiro Word localizados com sucesso."
                )

            else:
                resultado = localizar_arquivos_conversao(pasta)

                self.entry_arquivo_1.configure(state="normal")
                self.entry_arquivo_1.delete(0, "end")
                self.entry_arquivo_1.insert(0, resultado["arquivo_word"])
                self.entry_arquivo_1.configure(state="readonly")

                messagebox.showinfo(
                    "Arquivo encontrado",
                    "Primeiro Word encontrado com sucesso."
                )

        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def gerar_documentos_ui(self):
        if not self.pasta_var.get().strip():
            messagebox.showwarning(
                "Atenção",
                "Selecione uma pasta antes de executar."
            )
            return

        self.btn_gerar.configure(state="disabled")
        self.btn_localizar.configure(state="disabled")
        self.btn_criar.configure(state="disabled")
        self.btn_pasta.configure(state="disabled")

        threading.Thread(target=self._executar_com_loading, daemon=True).start()

    def _executar_com_loading(self):
        loading_proc = None

        try:
            pasta = self.normalizar_pasta_raiz(self.pasta_var.get().strip())
            self.pasta_var.set(pasta)

            caminho_script = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "wallpaper_loading_qt.py"
            )

            caminho_imagem = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "IMG_8286.PNG"
            )

            loading_proc = subprocess.Popen([
                sys.executable,
                caminho_script,
                caminho_imagem,
                "60000"
            ])

            modo = self.modo_var.get()

            if modo == "modelo":
                resultado = executar_processamento_modelo(
                    pasta_raiz=pasta,
                    quantidade_campos=self.quantidade_campos_var.get(),
                    gerar_pdf=self.gerar_pdf_var.get()
                )

                self.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Concluído",
                        f"Processamento finalizado com sucesso.\n\n"
                        f"Word gerados: {resultado['total_word']}\n"
                        f"PDF gerados: {resultado['total_pdf']}"
                    )
                )

            else:
                resultado = executar_conversao_word_para_pdf(
                    pasta_raiz=pasta
                )

                self.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Concluído",
                        f"Conversão finalizada com sucesso.\n\n"
                        f"Words lidos: {resultado['total_word_entrada']}\n"
                        f"PDFs gerados: {resultado['total_pdf']}"
                    )
                )

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Erro", str(e)))

        finally:
            if loading_proc is not None:
                try:
                    loading_proc.kill()
                except Exception:
                    pass

            self.after(0, self._reabilitar_botoes)

    def _reabilitar_botoes(self):
        self.btn_gerar.configure(state="normal")
        self.btn_localizar.configure(state="normal")
        self.btn_criar.configure(state="normal")
        self.btn_pasta.configure(state="normal")

    def _montar_layout(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, corner_radius=0, height=70)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        header.grid_propagate(False)

        titulo = ctk.CTkLabel(
            header,
            text="Gerador de Word e PDF",
            font=ctk.CTkFont(size=22, weight="bold")
        )
        titulo.grid(row=0, column=0, padx=16, pady=(12, 2), sticky="w")

        subtitulo = ctk.CTkLabel(
            header,
            text="Selecione a pasta, defina o modo e execute.",
            font=ctk.CTkFont(size=12)
        )
        subtitulo.grid(row=1, column=0, padx=16, pady=(0, 8), sticky="w")

        card = ctk.CTkFrame(self, corner_radius=14)
        card.grid(row=1, column=0, sticky="nsew", padx=14, pady=14)
        card.grid_columnconfigure(0, weight=1)

        self.lbl_pasta = ctk.CTkLabel(card, text="Pasta principal")
        self.lbl_pasta.grid(row=0, column=0, padx=14, pady=(12, 2), sticky="w")

        self.entry_pasta = ctk.CTkEntry(
            card,
            textvariable=self.pasta_var,
            height=34,
            placeholder_text="Selecione a pasta principal"
        )
        self.entry_pasta.grid(row=1, column=0, padx=14, pady=(0, 6), sticky="ew")

        self.frame_botoes_pasta = ctk.CTkFrame(card, fg_color="transparent")
        self.frame_botoes_pasta.grid(row=2, column=0, padx=14, pady=(0, 10), sticky="ew")
        self.frame_botoes_pasta.grid_columnconfigure((0, 1), weight=1)

        self.btn_pasta = ctk.CTkButton(
            self.frame_botoes_pasta,
            text="Selecionar pasta",
            height=34,
            command=self.escolher_pasta
        )
        self.btn_pasta.grid(row=0, column=0, padx=(0, 4), sticky="ew")

        self.btn_criar = ctk.CTkButton(
            self.frame_botoes_pasta,
            text="Criar estrutura",
            height=34,
            command=self.criar_pastas
        )
        self.btn_criar.grid(row=0, column=1, padx=(4, 0), sticky="ew")

        ctk.CTkLabel(card, text="Modo").grid(
            row=3, column=0, padx=14, pady=(4, 2), sticky="w"
        )

        self.segmentado_modo = ctk.CTkSegmentedButton(
            card,
            values=["modelo", "conversao"],
            variable=self.modo_var,
            command=lambda _: self.atualizar_modo(),
            height=32
        )
        self.segmentado_modo.grid(row=4, column=0, padx=14, pady=(0, 8), sticky="ew")
        self.segmentado_modo.set("modelo")

        self.frame_parametros = ctk.CTkFrame(card, fg_color="transparent")
        self.frame_parametros.grid(row=5, column=0, sticky="ew")
        self.frame_parametros.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self.frame_parametros, text="Parâmetros").grid(
            row=0, column=0, padx=14, pady=(0, 2), sticky="w"
        )

        self.slider_param = ctk.CTkSlider(
            self.frame_parametros,
            from_=1,
            to=30,
            number_of_steps=29,
            command=self.on_slider_change,
            height=14
        )
        self.slider_param.set(3)
        self.slider_param.grid(row=1, column=0, padx=14, pady=(0, 4), sticky="ew")

        self.lbl_qtd = ctk.CTkLabel(
            self.frame_parametros,
            text="3 parâmetros",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.lbl_qtd.grid(row=2, column=0, padx=14, pady=(0, 6), sticky="w")

        self.chk_pdf = ctk.CTkCheckBox(
            card,
            text="Gerar PDF",
            variable=self.gerar_pdf_var
        )
        self.chk_pdf.grid(row=6, column=0, padx=14, pady=(0, 8), sticky="w")

        self.lbl_arquivo_1 = ctk.CTkLabel(card, text="Primeiro Excel encontrado")
        self.lbl_arquivo_1.grid(row=7, column=0, padx=14, pady=(0, 2), sticky="w")

        self.entry_arquivo_1 = ctk.CTkEntry(card, height=32, state="readonly")
        self.entry_arquivo_1.grid(row=8, column=0, padx=14, pady=(0, 6), sticky="ew")

        self.lbl_arquivo_2 = ctk.CTkLabel(card, text="Primeiro Word da base encontrado")
        self.lbl_arquivo_2.grid(row=9, column=0, padx=14, pady=(0, 2), sticky="w")

        self.entry_arquivo_2 = ctk.CTkEntry(card, height=32, state="readonly")
        self.entry_arquivo_2.grid(row=10, column=0, padx=14, pady=(0, 10), sticky="ew")

        linha_botoes2 = ctk.CTkFrame(card, fg_color="transparent")
        linha_botoes2.grid(row=11, column=0, padx=14, pady=(0, 12), sticky="ew")
        linha_botoes2.grid_columnconfigure((0, 1), weight=1)

        self.btn_localizar = ctk.CTkButton(
            linha_botoes2,
            text="Localizar",
            height=36,
            command=self.localizar_arquivos_ui
        )
        self.btn_localizar.grid(row=0, column=0, padx=(0, 4), sticky="ew")

        self.btn_gerar = ctk.CTkButton(
            linha_botoes2,
            text="Executar",
            height=36,
            command=self.gerar_documentos_ui
        )
        self.btn_gerar.grid(row=0, column=1, padx=(4, 0), sticky="ew")


if __name__ == "__main__":
    app = App()
    app.mainloop()