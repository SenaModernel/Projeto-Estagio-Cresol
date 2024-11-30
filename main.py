import os
import pandas as pd
from tkinter import Tk, Button, Label, Listbox, filedialog, Menu, messagebox, ttk
from tkinter.simpledialog import askstring


class PlanilhaUnirApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Editor de Planilhas")
        self.root.geometry("1000x600")

        # Lista para armazenar caminhos de arquivos e planilhas
        self.arquivos = []
        self.planilhas = []
        self.planilha_atual = None
        self.coluna_selecionada = None  # Índice da coluna selecionada

        # Criar os painéis
        self.painel_esquerdo()
        self.painel_direito()

    def painel_esquerdo(self):
        """Configura o painel esquerdo com a lista de planilhas."""
        self.frame_esquerdo = ttk.Frame(self.root, width=200)
        self.frame_esquerdo.pack(side="left", fill="y")

        # Lista de planilhas carregadas
        self.label_planilhas = Label(self.frame_esquerdo, text="Planilhas Carregadas")
        self.label_planilhas.pack(pady=10)

        self.listbox = Listbox(self.frame_esquerdo, height=30, width=25)
        self.listbox.pack(padx=10, pady=5)
        self.listbox.bind("<<ListboxSelect>>", self.carregar_planilha)

        self.botao_add = Button(self.frame_esquerdo, text="Adicionar Planilha", command=self.adicionar_arquivo)
        self.botao_add.pack(pady=5)

        self.botao_remover = Button(self.frame_esquerdo, text="Remover Selecionada", command=self.remover_planilha)
        self.botao_remover.pack(pady=5)

    def painel_direito(self):
        """Configura o painel direito para exibir a planilha selecionada."""
        self.frame_direito = ttk.Frame(self.root)
        self.frame_direito.pack(side="right", fill="both", expand=True)

        self.label_planilha_atual = Label(self.frame_direito, text="Nenhuma Planilha Selecionada", font=("Arial", 14))
        self.label_planilha_atual.pack(pady=10)

        self.tree = ttk.Treeview(self.frame_direito)
        self.tree.pack(fill="both", expand=True)

        # Adicionar menu de contexto ao painel direito
        self.tree.bind("<Button-3>", self.menu_contexto)

    def adicionar_arquivo(self):
        """Adiciona uma nova planilha ao painel esquerdo."""
        caminho = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if caminho:
            try:
                planilha = pd.read_excel(caminho)
                self.arquivos.append(caminho)
                self.planilhas.append(planilha)
                self.listbox.insert("end", os.path.basename(caminho))
                messagebox.showinfo("Sucesso", "Planilha adicionada com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar o arquivo: {e}")

    def remover_planilha(self):
        """Remove a planilha selecionada."""
        try:
            index = self.listbox.curselection()[0]
            self.arquivos.pop(index)
            self.planilhas.pop(index)
            self.listbox.delete(index)

            # Limpar painel direito se a planilha exibida for removida
            self.planilha_atual = None
            self.label_planilha_atual.config(text="Nenhuma Planilha Selecionada")
            for item in self.tree.get_children():
                self.tree.delete(item)

            messagebox.showinfo("Sucesso", "Planilha removida com sucesso!")
        except IndexError:
            messagebox.showwarning("Aviso", "Nenhuma planilha selecionada para remover.")

    def carregar_planilha(self, event):
        """Carrega e exibe a planilha selecionada no painel direito."""
        try:
            index = self.listbox.curselection()[0]
            self.planilha_atual = self.planilhas[index]
            self.label_planilha_atual.config(text=f"Planilha: {os.path.basename(self.arquivos[index])}")

            # Limpar o Treeview antes de carregar novos dados
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Configurar as colunas do Treeview
            self.tree["columns"] = list(self.planilha_atual.columns)
            self.tree["show"] = "headings"

            for col in self.planilha_atual.columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, anchor="center")

            # Inserir os dados no Treeview
            for _, row in self.planilha_atual.iterrows():
                self.tree.insert("", "end", values=list(row))

        except IndexError:
            pass

    def menu_contexto(self, event):
        """Exibe o menu de contexto ao clicar com o botão direito."""
        # Verificar se há uma planilha carregada
        if self.planilha_atual is None:
            messagebox.showwarning("Aviso", "Nenhuma planilha selecionada para editar.")
            return

        # Identificar a coluna clicada
        coluna = self.tree.identify_column(event.x)
        if not coluna:
            messagebox.showwarning("Aviso", "Nenhuma coluna identificada.")
            return

        # Armazenar o índice da coluna clicada
        self.coluna_selecionada = int(coluna[1:]) - 1

        # Criar menu de contexto
        menu = Menu(self.root, tearoff=0)
        menu.add_command(label="Renomear Coluna", command=self.renomear_coluna)
        menu.add_command(label="Excluir Coluna", command=self.excluir_coluna)
        menu.post(event.x_root, event.y_root)

    def renomear_coluna(self):
        """Permite renomear uma coluna."""
        if self.planilha_atual is None:
            messagebox.showwarning("Aviso", "Nenhuma planilha selecionada.")
            return

        if self.coluna_selecionada is None:
            messagebox.showwarning("Aviso", "Nenhuma coluna selecionada.")
            return

        # Identificar o nome da coluna selecionada
        nome_coluna = self.tree["columns"][self.coluna_selecionada]

        # Solicitar o novo nome
        novo_nome = askstring("Renomear Coluna", f"Digite o novo nome para a coluna '{nome_coluna}':")
        if novo_nome:
            self.planilha_atual.rename(columns={nome_coluna: novo_nome}, inplace=True)
            self.carregar_planilha(None)

    def excluir_coluna(self):
        """Exclui a coluna selecionada."""
        if self.planilha_atual is None:
            messagebox.showwarning("Aviso", "Nenhuma planilha selecionada.")
            return

        if self.coluna_selecionada is None:
            messagebox.showwarning("Aviso", "Nenhuma coluna selecionada.")
            return

        # Identificar o nome da coluna selecionada
        nome_coluna = self.tree["columns"][self.coluna_selecionada]

        # Confirmar exclusão
        confirmacao = messagebox.askyesno("Confirmação", f"Deseja realmente excluir a coluna '{nome_coluna}'?")
        if confirmacao:
            self.planilha_atual.drop(columns=[nome_coluna], inplace=True)
            self.carregar_planilha(None)


if __name__ == "__main__":
    root = Tk()
    app = PlanilhaUnirApp(root)
    root.mainloop()
