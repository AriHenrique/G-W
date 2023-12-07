import tkinter as tk
from tkinter import filedialog, messagebox
import pyodbc
import pandas as pd


class InterfaceGrafica:
    def __init__(self, root):
        self.root = root
        self.root.title("Carregar/Salvar Dados")
        self.root.geometry("400x200")

        self.label = tk.Label(root, text="Escolha uma opção:")
        self.label.pack(pady=10)

        self.carregar_botao = tk.Button(root, text="Carregar a PLANILHA para o Banco de Dados", command=self.carregar_para_banco)
        self.carregar_botao.pack(pady=10)

        self.salvar_botao = tk.Button(root, text="Salvar Dados do Banco para a PLANILHA", command=self.salvar_para_xlsx)
        self.salvar_botao.pack(pady=10)

        self.mensagem_var = tk.StringVar()
        self.mensagem_label = tk.Label(root, textvariable=self.mensagem_var, fg="red")
        self.mensagem_label.pack(pady=10)

    def carregar_para_banco(self):
        arquivo_xlsx = filedialog.askopenfilename(title="Selecione o arquivo XLSX", filetypes=[("Arquivos XLSX", "*.xlsx")])
        if arquivo_xlsx:
            self.mensagem_var.set(f"Carregando dados de {arquivo_xlsx} para o banco de dados...")
            try:
                carregar_dados_para_banco(arquivo_xlsx)
                self.mensagem_var.set("Atualização realizada com sucesso!")
            except ValueError as e:
                self.exibir_erro(e)

    def salvar_para_xlsx(self):
        pasta_destino = filedialog.askdirectory(title="Selecione a pasta de destino")
        if pasta_destino:
            self.mensagem_var.set(f"Salvando dados do banco de dados para {pasta_destino}...")
            try:
                carregar_dados_para_planilha(pasta_destino)
                self.mensagem_var.set(f'Dados extraídos para {pasta_destino}/SERVICOS.xlsx')
            except ValueError as e:
                self.exibir_erro(e)

    def exibir_erro(self, mensagem):
        messagebox.showerror("Erro", mensagem)
        self.mensagem_var.set("Erro: " + mensagem)


def carregar_dados_para_banco(arquivo_xlsx):
    conn_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\SHARMAQ\SHOficina\dados.mdb;'
    conn = pyodbc.connect(conn_string)
    cursor = conn.cursor()

    dfs = pd.read_excel(arquivo_xlsx, sheet_name=None)

    for sheet_name, df in dfs.items():
        for index, row in df.iterrows():
            check_query = "SELECT COUNT(*) FROM SERVICOS WHERE DESCRICAO = ? AND GRUPO = ?"
            check_values = (row['DESCRICAO'], sheet_name)
            cursor.execute(check_query, check_values)
            row_count = cursor.fetchone()[0]

            number_list = ['VALOR', 'COMISSAO', 'CUSTO_UNT']
            for i, name in enumerate(number_list):
                try:
                    float(row[name])
                except ValueError:
                    raise ValueError(f"Você digitou um valor incorreto na folha {sheet_name} no campo {number_list[i]}."
                                     f"Apenas números inteiros ou separados por vírgula são aceitos.")
                    exit()
            if row_count:
                update_query = "UPDATE SERVICOS SET VALOR = ?, COMISSAO = ?, CUSTO_UNT = ? WHERE DESCRICAO = ? AND GRUPO = ?"
                update_values = (float(row['VALOR'])
                                 , float(row['COMISSAO'])
                                 , float(row['CUSTO_UNT'])
                                 , row['DESCRICAO']
                                 , sheet_name
                                 )
                cursor.execute(update_query, update_values)
            else:
                insert_query = "INSERT INTO SERVICOS (DESCRICAO, VALOR, COMISSAO, GRUPO, CUSTO_UNT) VALUES (?, ?, ?, ?, ?)"
                insert_values = (
                    row['DESCRICAO']
                    , float(row['VALOR'])
                    , float(row['COMISSAO'])
                    , sheet_name
                    , float(row['CUSTO_UNT'])
                )
                cursor.execute(insert_query, insert_values)

    conn.commit()
    cursor.execute('SELECT * FROM SERVICOS')
    conn.close()


def carregar_dados_para_planilha(pasta_destino):
    conn_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\SHARMAQ\SHOficina\dados.mdb;'
    conn = pyodbc.connect(conn_string)
    cursor = conn.cursor()

    select_query = 'SELECT * FROM SERVICOS'
    df = pd.read_sql(select_query, conn)
    columns_to_remove = ['CODIGO', 'LCP116']
    df = df.drop(columns=columns_to_remove, errors='ignore')
    output_excel_path = f'{pasta_destino}/SERVICOS.xlsx'
    grouped_dfs = dict(tuple(df.groupby('GRUPO')))

    with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
        for group_name, group_df in grouped_dfs.items():
            group_df = group_df.drop(columns=['GRUPO'], errors='ignore')
            group_df.to_excel(writer, sheet_name=group_name, index=False)

    conn.close()


if __name__ == "__main__":
    root = tk.Tk()
    app = InterfaceGrafica(root)
    root.mainloop()
