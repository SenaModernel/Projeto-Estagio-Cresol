import os
import pandas as pd
from tkinter import Tk, Listbox, Button, Label, messagebox, filedialog


class PlanilhaUnirApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Unir Planilhas")
        self.root.geometry("400x350")

        # Lista para armazenar caminhos de arquivos
        self.arquivos = []

        # janelas
        self.label = Label(root, text="Selecione os arquivos para unir:")
        self.label.pack(pady=10)

        self.listbox = Listbox(root, height=10, width=50)
        self.listbox.pack(padx=10, pady=5)

        self.botao_add = Button(root, text="Adicionar Arquivo", command=self.adicionar_arquivo)
        self.botao_add.pack(pady=5)

        self.botao_remover = Button(root, text="Remover Selecionado", command=self.remover_selecionado)
        self.botao_remover.pack(pady=5)

        self.botao_unir = Button(root, text="Unir Planilhas", command=self.unir_planilhas, state="disabled")
        self.botao_unir.pack(pady=10)

    def adicionar_arquivo(self):
        """Adiciona um arquivo à lista."""
        caminho = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if caminho:
            self.arquivos.append(caminho)
            # Adicionar somente o nome do arquivo na Listbox
            self.listbox.insert("end", os.path.basename(caminho))
            # Habilitar o botão "Unir" se houver arquivos selecionados
            self.botao_unir.config(state="normal")

    def remover_selecionado(self):
        """Remove o arquivo selecionado na lista."""
        try:
            index = self.listbox.curselection()[0]
            self.arquivos.pop(index)
            self.listbox.delete(index)

            # Desabilitar o botão "Unir" se não houver mais arquivos
            if not self.arquivos:
                self.botao_unir.config(state="disabled")
        except IndexError:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado para remover.")

    def unir_planilhas(self):
        """Une as planilhas selecionadas."""
        if not self.arquivos:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado para unir.")
            return

        try:
            # Carregar todas as planilhas
            planilhas = [pd.read_excel(caminho) for caminho in self.arquivos]

            # Validar as planilhas
            colunas_base = planilhas[0].columns
            for idx, planilha in enumerate(planilhas[1:], start=2):
                if planilha.shape[1] != len(colunas_base):
                    raise ValueError(f"A planilha {idx} possui número de colunas diferente.")
                if not all(planilha.columns == colunas_base):
                    raise ValueError(f"A planilha {idx} possui nomes de colunas diferentes.")

            # Concatenar as planilhas
            resultado = pd.concat(planilhas, ignore_index=True)

            # Salvar a planilha final
            caminho_saida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if caminho_saida:
                resultado.to_excel(caminho_saida, index=False)
                messagebox.showinfo("Sucesso", f"Planilhas unidas com sucesso!\nSalvo em: {caminho_saida}")
            else:
                messagebox.showwarning("Aviso", "Nenhum local selecionado para salvar o arquivo.")

        except ValueError as e:
            messagebox.showerror("Erro de Validação", str(e))
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro inesperado: {e}")


if __name__ == "__main__":
    root = Tk()
    app = PlanilhaUnirApp(root)
    root.mainloop()
