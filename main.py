import os
import pandas as pd
from tkinter import Tk, Listbox, Button, Label, messagebox, filedialog, Toplevel, Entry, ttk


class PlanilhaUnirApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Unir Planilhas")
        self.root.geometry("500x500")

        # Lista para armazenar caminhos de arquivos e planilhas
        self.arquivos = []
        self.planilhas = []
        self.erros_colunas = False

        # Widgets
        self.label = Label(root, text="Selecione os arquivos para unir:")
        self.label.pack(pady=10)

        self.listbox = Listbox(root, height=15, width=60)
        self.listbox.pack(padx=10, pady=5)

        self.botao_add = Button(root, text="Adicionar Arquivo", command=self.adicionar_arquivo)
        self.botao_add.pack(pady=5)

        self.botao_remover = Button(root, text="Remover Selecionado", command=self.remover_selecionado)
        self.botao_remover.pack(pady=5)

        self.botao_ajustar = Button(root, text="Resolver Erros", command=self.ajustar_colunas, state="disabled")
        self.botao_ajustar.pack(pady=5)

        self.botao_unir = Button(root, text="Unir Planilhas", command=self.unir_planilhas, state="disabled")
        self.botao_unir.pack(pady=10)

        self.botao_encerrar = Button(root, text="Encerrar Aplicação", command=self.encerrar_aplicacao)
        self.botao_encerrar.pack(pady=10)

    def adicionar_arquivo(self):
        """Adiciona um arquivo à lista e carrega a planilha."""
        caminho = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if caminho:
            try:
                planilha = pd.read_excel(caminho)
                self.arquivos.append(caminho)
                self.planilhas.append(planilha)
                self.listbox.insert("end", os.path.basename(caminho))

                # Após adicionar, verificar se há erros
                self.verificar_erros()

            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar o arquivo: {e}")

    def remover_selecionado(self):
        """Remove o arquivo selecionado na lista."""
        try:
            index = self.listbox.curselection()[0]
            self.arquivos.pop(index)
            self.planilhas.pop(index)
            self.listbox.delete(index)

            # Após remover, verificar novamente se há erros
            self.verificar_erros()

        except IndexError:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado para remover.")

    def verificar_erros(self):
        """Verifica se existem erros de colunas entre as planilhas."""
        if len(self.planilhas) < 2:
            # Desabilitar todos os botões se não houver planilhas suficientes
            self.botao_unir.config(state="disabled")
            self.botao_ajustar.config(state="disabled")
            return

        # Comparar colunas
        colunas_base = self.planilhas[0].columns
        self.erros_colunas = False
        for planilha in self.planilhas[1:]:
            if not planilha.columns.equals(colunas_base):
                self.erros_colunas = True
                break

        # Habilitar os botões com base no resultado
        if self.erros_colunas:
            self.botao_ajustar.config(state="normal")
            self.botao_unir.config(state="disabled")
        else:
            self.botao_ajustar.config(state="disabled")
            self.botao_unir.config(state="normal")

    def ajustar_colunas(self):
        """Exibe a planilha com erros e permite ajustes manuais."""
        if not self.erros_colunas:
            messagebox.showinfo("Sem Erros", "Nenhum erro encontrado nas colunas.")
            return

        for idx, planilha in enumerate(self.planilhas, start=1):
            self.exibir_planilha(idx, planilha)

    def exibir_planilha(self, index, planilha):
        """Exibe a planilha em uma janela visual."""
        ajuste_window = Toplevel(self.root)
        ajuste_window.title(f"Editar Planilha {index}")
        ajuste_window.geometry("700x500")

        # Exibir colunas como Treeview
        tree = ttk.Treeview(ajuste_window)
        tree.pack(fill="both", expand=True)

        tree["columns"] = list(planilha.columns)
        tree["show"] = "headings"

        for col in planilha.columns:
            tree.heading(col, text=col, command=lambda _col=col: self.alterar_nome_coluna(index, planilha, _col))
            tree.column(col, width=100, anchor="center")

        # Inserir dados na Treeview
        for i, row in planilha.iterrows():
            tree.insert("", "end", values=list(row))

        # Botões para salvar ou descartar alterações
        botao_salvar = Button(ajuste_window, text="Salvar Alterações", command=lambda: self.salvar_alteracoes(index, planilha))
        botao_salvar.pack(pady=5)

        botao_descartar = Button(ajuste_window, text="Descartar Alterações", command=ajuste_window.destroy)
        botao_descartar.pack(pady=5)

    def alterar_nome_coluna(self, index, planilha, coluna_atual):
        """Permite alterar o nome de uma coluna."""
        novo_nome = askstring("Renomear Coluna", f"Digite o novo nome para a coluna '{coluna_atual}':")
        if novo_nome:
            planilha.rename(columns={coluna_atual: novo_nome}, inplace=True)
            messagebox.showinfo("Alteração Realizada", f"A coluna '{coluna_atual}' foi renomeada para '{novo_nome}'.")

    def salvar_alteracoes(self, index, planilha):
        """Salva alterações e revalida planilhas."""
        self.planilhas[index - 1] = planilha
        messagebox.showinfo("Sucesso", f"As alterações na planilha {index} foram salvas.")
        self.verificar_erros()

    def unir_planilhas(self):
        """Une as planilhas selecionadas."""
        try:
            resultado = pd.concat(self.planilhas, ignore_index=True)
            caminho_saida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if caminho_saida:
                resultado.to_excel(caminho_saida, index=False)
                messagebox.showinfo("Sucesso", f"Planilhas unidas com sucesso!\nSalvo em: {caminho_saida}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao unir planilhas: {e}")

    def encerrar_aplicacao(self):
        """Encerra a aplicação."""
        self.root.quit()
        self.root.destroy()


if __name__ == "__main__":
    root = Tk()
    app = PlanilhaUnirApp(root)
    root.mainloop()
