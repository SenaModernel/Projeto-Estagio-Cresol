import os
import pandas as pd
from tkinter import Tk, Button, Label, Listbox, filedialog, Menu, messagebox, ttk, Scrollbar
from tkinter.simpledialog import askstring


class PlanilhaUnirApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Editor de Planilhas")
        self.root.geometry("1000x600")

        # Configuração de estilo para o Treeview
        self.estilizar_treeview()

        # Lista para armazenar caminhos de arquivos e planilhas
        self.arquivos = []
        self.planilhas = []
        self.planilha_atual = None
        self.coluna_selecionada = None  # Índice da coluna selecionada

        # Criar os painéis
        self.painel_esquerdo()
        self.painel_direito()

    def estilizar_treeview(self):
        """Aplica estilos personalizados ao Treeview."""
        estilo = ttk.Style()
        estilo.theme_use("default")

        # Estilo do cabeçalho em negrito
        estilo.configure("Treeview.Heading", font=("Arial", 10, "bold"))

        # Estilo das células (com linhas de grade)
        estilo.configure(
            "Treeview",
            font=("Arial", 10),
            rowheight=25,
            background="#f9f9f9",
            fieldbackground="#ffffff",
            borderwidth=1
        )

        # Alternância de cores nas linhas
        estilo.map(
            "Treeview",
            background=[("selected", "#bfbfbf"), ("!selected", "#ffffff")],
            foreground=[("selected", "black"), ("!selected", "black")]
        )

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

        self.botao_unir = Button(self.frame_esquerdo, text="Unir Planilhas", command=self.unir_planilhas)
        self.botao_unir.pack(pady=10)

    def painel_direito(self):
        """Configura o painel direito para exibir a planilha selecionada."""
        self.frame_direito = ttk.Frame(self.root)
        self.frame_direito.pack(side="right", fill="both", expand=True)

        self.label_planilha_atual = Label(self.frame_direito, text="Nenhuma Planilha Selecionada", font=("Arial", 14))
        self.label_planilha_atual.pack(pady=10)

        # Frame para a tabela e as barras de rolagem
        self.frame_treeview = ttk.Frame(self.frame_direito)
        self.frame_treeview.pack(fill="both", expand=True)

        # Treeview com barras de rolagem
        self.tree = ttk.Treeview(self.frame_treeview, show="headings")
        self.scrollbar_vertical = Scrollbar(self.frame_treeview, orient="vertical", command=self.tree.yview)
        self.scrollbar_horizontal = Scrollbar(self.frame_treeview, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.scrollbar_vertical.set, xscrollcommand=self.scrollbar_horizontal.set)

        # Posicionar o Treeview e as barras de rolagem
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.scrollbar_vertical.grid(row=0, column=1, sticky="ns")
        self.scrollbar_horizontal.grid(row=1, column=0, sticky="ew")

        # Configurar expansão do Treeview dentro do frame
        self.frame_treeview.grid_rowconfigure(0, weight=1)
        self.frame_treeview.grid_columnconfigure(0, weight=1)

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
            self.tree.delete(*self.tree.get_children())

            # Configurar as colunas do Treeview
            self.tree["columns"] = list(self.planilha_atual.columns)
            self.tree["show"] = "headings"

            # Ajustar largura automática para as colunas
            for col in self.planilha_atual.columns:
                largura_cabecalho = len(col)
                largura_dados = self.planilha_atual[col].astype(str).map(len).max()
                largura_coluna = max(largura_cabecalho, largura_dados) * 10  # Fator de ajuste para largura
                self.tree.heading(col, text=col)  # Define o cabeçalho
                self.tree.column(col, anchor="center", width=largura_coluna)  # Ajusta a largura da coluna

            # Inserir os dados no Treeview
            for _, row in self.planilha_atual.iterrows():
                self.tree.insert("", "end", values=list(row))

        except IndexError:
            messagebox.showwarning("Aviso", "Nenhuma planilha selecionada.")

    def menu_contexto(self, event):
        """Exibe o menu de contexto ao clicar com o botão direito."""
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

        nome_coluna = self.tree["columns"][self.coluna_selecionada]

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

        nome_coluna = self.tree["columns"][self.coluna_selecionada]

        confirmacao = messagebox.askyesno("Confirmação", f"Deseja realmente excluir a coluna '{nome_coluna}'?")
        if confirmacao:
            self.planilha_atual.drop(columns=[nome_coluna], inplace=True)
            self.carregar_planilha(None)

    def unir_planilhas(self):
        """Une as planilhas carregadas."""
        if len(self.planilhas) < 2:
            messagebox.showwarning("Aviso", "Adicione pelo menos duas planilhas para unir.")
            return

        try:
            colunas_referencia = list(self.planilhas[0].columns)
            for planilha in self.planilhas:
                if list(planilha.columns) != colunas_referencia:
                    raise ValueError("As planilhas possuem colunas incompatíveis.")

            resultado = pd.concat(self.planilhas, ignore_index=True)

            caminho_saida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if caminho_saida:
                resultado.to_excel(caminho_saida, index=False)
                messagebox.showinfo("Sucesso", f"Planilhas unidas e salvas em: {caminho_saida}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao unir planilhas: {e}")


if __name__ == "__main__":
    root = Tk()
    app = PlanilhaUnirApp(root)
    root.mainloop()