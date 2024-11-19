import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename

def unir_planilhas():
    try:

        Tk().withdraw()  # Oculta a janela principal do Tkinter
        
        # Selecionar o primeiro arquivo
        print("Selecione a primeira planilha:")
        caminho_arquivo1 = askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not caminho_arquivo1:
            print("Nenhum arquivo selecionado.")
            return

        # Selecionar o segundo arquivo
        print("Selecione a segunda planilha:")
        caminho_arquivo2 = askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not caminho_arquivo2:
            print("Nenhum arquivo selecionado.")
            return

        # Carregar as planilhas
        planilha1 = pd.read_excel(caminho_arquivo1)
        planilha2 = pd.read_excel(caminho_arquivo2)

        # Verificar número de colunas
        if planilha1.shape[1] != planilha2.shape[1]:
            raise ValueError("As planilhas possuem números diferentes de colunas.")

        # Verificar nomes das colunas
        if not all(planilha1.columns == planilha2.columns):
            raise ValueError("Os nomes das colunas não são iguais.")

        # Concatenar as planilhas
        resultado = pd.concat([planilha1, planilha2], ignore_index=True)

        # Selecionar o local para salvar o arquivo final
        print("Selecione onde salvar a planilha resultante:")
        caminho_saida = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not caminho_saida:
            print("Nenhum local selecionado. Operação cancelada.")
            return

        # Salvar o arquivo unido
        resultado.to_excel(caminho_saida, index=False)
        print("Planilhas unidas com sucesso!")

    except Exception as e:
        print(f"Erro: {e}")

# Executar a função
unir_planilhas()
