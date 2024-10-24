import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import os

class AtualizacaoPrecosApp:
    def __init__(self, master):
        self.master = master
        master.title("Atualização de Preços")
        
        # Configurações de estilo
        master.configure(bg="#f7f7f7")
        self.font = ("Helvetica", 12)

        # Frame para centralizar o conteúdo
        self.frame = tk.Frame(master, bg="#f7f7f7")
        self.frame.pack(expand=True)

        # Label e campo de texto para exibir o caminho do arquivo do catálogo
        self.label_catalogo = tk.Label(self.frame, text="Selecione o arquivo do catálogo:", bg="#f7f7f7", font=self.font)
        self.label_catalogo.pack(pady=10)

        self.caminho_catalogo = tk.Entry(self.frame, width=50, font=self.font)
        self.caminho_catalogo.pack(pady=5)

        # Botão para carregar o catálogo
        self.botao_carregar_catalogo = self.criar_botao(self.frame, "Carregar Catálogo", self.carregar_catalogo, "#a8dadc")
        self.botao_carregar_catalogo.pack(pady=10)

        # Label e campo de texto para exibir o caminho do arquivo da loja
        self.label_loja = tk.Label(self.frame, text="Selecione o arquivo da loja:", bg="#f7f7f7", font=self.font)
        self.label_loja.pack(pady=10)

        self.caminho_loja = tk.Entry(self.frame, width=50, font=self.font)
        self.caminho_loja.pack(pady=5)

        # Botão para carregar a loja
        self.botao_carregar_loja = self.criar_botao(self.frame, "Carregar Loja", self.carregar_loja, "#a8dadc")
        self.botao_carregar_loja.pack(pady=10)

        # Botão para iniciar a atualização
        self.botao_atualizar = self.criar_botao(self.frame, "Comparar Preços", self.comparar_precos, "#ffabab")
        self.botao_atualizar.pack(pady=10)

        # Área de texto para exibir mensagens
        self.mensagem = tk.Text(self.frame, height=10, width=50, font=self.font, bg="#ffffff", wrap=tk.WORD)
        self.mensagem.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Adicionando um scrollbar à área de texto
        self.scrollbar = tk.Scrollbar(self.frame, command=self.mensagem.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.mensagem['yscrollcommand'] = self.scrollbar.set

        self.catalogo_df = None
        self.loja_df = None

    def criar_botao(self, parent, texto, comando, cor):
        # Cria um botão arredondado usando um Canvas
        canvas = tk.Canvas(parent, width=200, height=40, bg=cor, highlightthickness=0, borderwidth=0)
        canvas.create_arc(0, 0, 40, 40, start=90, extent=90, fill=cor, outline="")
        canvas.create_arc(160, 0, 200, 40, start=0, extent=90, fill=cor, outline="")
        canvas.create_rectangle(20, 0, 180, 40, fill=cor, outline="")
        canvas.create_text(100, 20, text=texto, fill="black", font=self.font)
        canvas.bind("<Button-1>", lambda event: comando())
        return canvas

    def carregar_catalogo(self):
        # Método para carregar o arquivo do catálogo
        arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if arquivo:
            self.caminho_catalogo.delete(0, tk.END)  # Limpa o campo de texto
            self.caminho_catalogo.insert(0, arquivo)  # Insere o caminho do arquivo
            # Lê o arquivo do catálogo
            if arquivo.endswith('.xlsx'):
                self.catalogo_df = pd.read_excel(arquivo)
            else:
                self.catalogo_df = pd.read_csv(arquivo)

            # Verifica os nomes das colunas
            self.catalogo_df.columns = self.catalogo_df.columns.str.strip().str.lower()  # Normaliza os nomes das colunas
            print(self.catalogo_df.columns)  # Debug: imprime os nomes das colunas
            self.mensagem.insert(tk.END, "Catálogo carregado com sucesso!\n")

    def carregar_loja(self):
    # Método para carregar o arquivo da loja
        arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if arquivo:
            self.caminho_loja.delete(0, tk.END)  # Limpa o campo de texto
            self.caminho_loja.insert(0, arquivo)  # Insere o caminho do arquivo
            # Lê o arquivo da loja
            if arquivo.endswith('.xlsx'):
                self.loja_df = pd.read_excel(arquivo)
            else:
                self.loja_df = pd.read_csv(arquivo)

            # Verifica os nomes das colunas
            self.loja_df.columns = self.loja_df.columns.str.strip().str.lower()  # Normaliza os nomes das colunas
            print(self.loja_df.columns)  # Debug: imprime os nomes das colunas
            self.mensagem.insert(tk.END, "Loja carregada com sucesso!\n")

    def comparar_precos(self):
        # Método para comparar os preços entre o catálogo e a loja
        if self.catalogo_df is None or self.loja_df is None:
            self.mensagem.insert(tk.END, "Por favor, carregue ambos os arquivos (catálogo e loja).\n")
            return

        # Normaliza os preços para garantir que sejam numéricos
        self.catalogo_df['preco'] = pd.to_numeric(self.catalogo_df['preco'], errors='coerce')
        self.loja_df['preco'] = pd.to_numeric(self.loja_df['preco'], errors='coerce')

        produtos_defasados = []

        for index, row in self.catalogo_df.iterrows():
            produto_catalogo = row['produto']
            preco_catalogo = row['preco']

            # Verifica se o produto está na loja
            loja_produto = self.loja_df[self.loja_df['produto'] == produto_catalogo]
            if not loja_produto.empty:
                preco_loja = loja_produto['preco'].values[0]
                print(f"Comparando: {produto_catalogo} - Preço Catálogo: {preco_catalogo}, Preço Loja: {preco_loja}")  # Debug
                
                # Adiciona produtos com preços defasados
                if preco_loja < preco_catalogo:
                    produtos_defasados.append((produto_catalogo, preco_loja, preco_catalogo, "menor"))
                elif preco_loja > preco_catalogo:
                    produtos_defasados.append((produto_catalogo, preco_loja, preco_catalogo, "maior"))

        # Exibe os produtos com preços defasados
        if produtos_defasados:
            self.mensagem.insert(tk.END, "Produtos com preços defasados:\n")
            for produto, preco_loja, preco_catalogo, tipo in produtos_defasados:
                if tipo == "menor":
                    self.mensagem.insert(tk.END, f"{produto}: Preço na loja: {preco_loja}, Preço no catálogo: {preco_catalogo} (menor na loja)\n")
                else:
                    self.mensagem.insert(tk.END, f"{produto}: Preço na loja: {preco_loja}, Preço no catálogo: {preco_catalogo} (maior na loja)\n")
            
           # Adiciona o botão "Atualizar"
            self.botao_atualizar = tk.Button(self.master, text="Atualizar", bg="#ffabab", width=25, height=2, font=("Helvetica", 12), command=self.gerar_catalogo_atual)  # Corrigido aqui
            self.botao_atualizar.pack()  # Adiciona o botão à interface
        else:
            self.mensagem.insert(tk.END, "Todos os preços estão atualizados.\n")

    def gerar_catalogo_atual(self):
        # Verifica se o catálogo foi carregado
        if self.catalogo_df is None:
            self.mensagem.insert(tk.END, "Nenhum catálogo carregado para atualizar.\n")
            return

        # Gera o nome do arquivo com data e hora
        data_hora = datetime.now().strftime("%d%m%Y_%H%M")
        nome_arquivo = f"Catalogo_atual_{data_hora}.xlsx"
        caminho = os.path.join("C:\\Users\\pedro\\Documents\\Python\\teste\\automacaofornecedores\\Catalogos_atualizados", nome_arquivo)

        # Salva o DataFrame como XLSX
        self.catalogo_df.to_excel(caminho, index=False)
        self.mensagem.insert(tk.END, f"Catálogo atualizado salvo em: {caminho}\n")
        

if __name__ == "__main__":
    root = tk.Tk()
    app = AtualizacaoPrecosApp(root)
    root.mainloop()