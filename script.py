import pandas as pd
import pyodbc
import tkinter as tk
from tkinter import Label, Entry, Button, filedialog, messagebox
import os
from datetime import datetime, timedelta

def gerar_excel(df_resultado):
    # Define o nome do arquivo com a data do dia anterior
    data_anterior = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    output_path = os.path.join(os.getcwd(), f"ata_refeitorio_{data_anterior}.xlsx")
    try:
        df_resultado.to_excel(output_path, index=False)
        messagebox.showinfo("Sucesso", f"Arquivo Excel gerado com sucesso: {output_path}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar o arquivo Excel: {e}")

def main():
    # Configuração da janela principal
    root = tk.Tk()
    root.title("Geração de Relatório do Refeitório")
    root.geometry("800x500")

    # Labels e entradas com valores pré-preenchidos para conexão
    Label(root, text="Servidor:", font=("Arial", 10)).pack(pady=5)
    server_entry = Entry(root, font=("Arial", 10))
    server_entry.pack(pady=5)

    Label(root, text="Banco de Dados:", font=("Arial", 10)).pack(pady=5)
    database_entry = Entry(root, font=("Arial", 10))
    database_entry.pack(pady=5)

    Label(root, text="Usuário:", font=("Arial", 10)).pack(pady=5)
    username_entry = Entry(root, font=("Arial", 10))
    username_entry.pack(pady=5)

    Label(root, text="Senha:", font=("Arial", 10)).pack(pady=5)
    password_entry = Entry(root, font=("Arial", 10), show="*")
    password_entry.pack(pady=5)

    # Variáveis para armazenar os caminhos dos arquivos
    file_path1 = None
    file_path2 = None

    # Funções para selecionar os arquivos
    def selecionar_arquivo1():
        nonlocal file_path1
        file_path1 = filedialog.askopenfilename(title="Selecione o primeiro arquivo Excel", filetypes=[("Excel files", "*.xlsx")])
        if file_path1:
            file_label1.config(text=f"Primeiro arquivo: {os.path.basename(file_path1)}")

    def selecionar_arquivo2():
        nonlocal file_path2
        file_path2 = filedialog.askopenfilename(title="Selecione o segundo arquivo Excel", filetypes=[("Excel files", "*.xlsx")])
        if file_path2:
            file_label2.config(text=f"Segundo arquivo: {os.path.basename(file_path2)}")

    # Botões e Labels para selecionar os arquivos
    file_label1 = Label(root, text="Nenhum arquivo selecionado", font=("Arial", 10), fg="grey")
    file_label1.pack(pady=5)
    select_button1 = Button(root, text="Selecionar Primeiro Arquivo", font=("Arial", 10), command=selecionar_arquivo1)
    select_button1.pack(pady=5)

    file_label2 = Label(root, text="Nenhum arquivo selecionado", font=("Arial", 10), fg="grey")
    file_label2.pack(pady=5)
    select_button2 = Button(root, text="Selecionar Segundo Arquivo", font=("Arial", 10), command=selecionar_arquivo2)
    select_button2.pack(pady=5)

    # Função para conectar ao banco de dados e gerar o relatório
    def conectar_e_gerar():
        if not file_path1 or not file_path2:
            messagebox.showwarning("Atenção", "Por favor, selecione ambos os arquivos.")
            return

        # Lê os arquivos Excel selecionados
        try:
            df_excel1 = pd.read_excel(file_path1)
            df_excel2 = pd.read_excel(file_path2)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler os arquivos Excel: {e}")
            return

        # Verifica se a coluna 'Cod_Cracha' existe nos arquivos Excel
        if 'Cod_Cracha' not in df_excel1.columns or 'Cod_Cracha' not in df_excel2.columns:
            messagebox.showwarning("Atenção", "A coluna 'Cod_Cracha' não foi encontrada em um dos arquivos Excel.")
            return

        # Combina os dados das colunas 'Cod_Cracha' dos dois arquivos
        cod_cracha_combined = pd.concat([df_excel1['Cod_Cracha'], df_excel2['Cod_Cracha']]).unique()

        # Configurações de conexão
        server = server_entry.get()
        database = database_entry.get()
        username = username_entry.get()
        password = password_entry.get()

        try:
            conn = pyodbc.connect(
                f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
            )
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao conectar ao banco de dados: {e}")
            return

        # Executa a consulta no banco de dados
        try:
            query = """
                 SELECT TOP (3500) [BADGE]
                    ,[USERID]
                FROM [sqlexample].[dbo].[BadgeLookup]
               
                """
            df_db = pd.read_sql_query(query, conn)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao executar a consulta no banco de dados: {e}")
            return
        finally:
            conn.close()

        # Verifica se a coluna 'Badge' existe no banco de dados
        if 'BADGE' not in df_db.columns:
            messagebox.showwarning("Atenção", "A coluna 'Badge' não foi encontrada no banco de dados.")
            return

        # Filtra os registros onde 'Cod_Cracha' combinado corresponde ao 'Badge' do banco de dados
        df_resultado = df_db[df_db['BADGE'].isin(cod_cracha_combined)]

        # Gera Excel
        gerar_excel(df_resultado)

    # Gerar relatório
    gerar_button = Button(root, text="Gerar Relatório", font=("Arial", 12), command=conectar_e_gerar)
    gerar_button.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
