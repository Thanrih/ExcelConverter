import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


def selecionar_arquivo():
    """Abre um diálogo para o usuário selecionar o arquivo Excel"""
    global excel_file
    excel_file = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if excel_file:
        label_arquivo['text'] = f"Arquivo selecionado: {excel_file}"
    else:
        label_arquivo['text'] = "Nenhum arquivo selecionado."


def formatar_dados(df):
    #Formata os campos 'MES_ANO_DIREITO' (data) e 'CPF'
    # Formatar a coluna de data
    if 'MES_ANO_DIREITO' in df.columns:
        df['MES_ANO_DIREITO'] = pd.to_datetime(df['MES_ANO_DIREITO'], errors='coerce').dt.strftime('%d/%m/%Y')

    # Formatar o CPF com pontos e traço
    if 'CPF' in df.columns:
        df['CPF'] = df['CPF'].astype(str).str.zfill(11)  # Garante 11 dígitos
        df['CPF'] = df['CPF'].str.replace(
            r'(\d{3})(\d{3})(\d{3})(\d{2})', r'\1.\2.\3-\4', regex=True
        )

    return df


def processar_arquivo():
    """Processa o arquivo Excel e salva como TXT"""
    try:
        # Verifica se o arquivo foi selecionado
        if not excel_file:
            raise ValueError("Por favor, selecione um arquivo Excel.")

        # Carrega o arquivo Excel
        try:
            df = pd.read_excel(excel_file, engine='openpyxl')
        except:
            df = pd.read_excel(excel_file, engine='xlrd')

        # Obter colunas selecionadas
        colunas = entry_colunas.get().split(',')
        colunas = [col.strip() for col in colunas]  # Remove espaços extras

        # Filtrar as colunas
        df_filtrado = df[colunas]

        # Formatar os dados
        df_formatado = formatar_dados(df_filtrado)

        # Escolher onde salvar o arquivo
        output_file = filedialog.asksaveasfilename(
            title="Salvar arquivo como",
            defaultextension=".txt",
            filetypes=[("Arquivo de Texto", "*.txt")]
        )
        if not output_file:
            raise ValueError("Nenhum caminho para salvar foi selecionado.")

        # Salvar o arquivo filtrado em formato TXT
        df_formatado.to_csv(output_file, sep=';', index=False)

        # Sucesso
        messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em:\n{output_file}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")


# Criar a janela principal
janela = tk.Tk()
janela.title("Conversor de Excel para TXT")

# Variável global para armazenar o arquivo selecionado
excel_file = None

# Elementos da interface
label_titulo = tk.Label(janela, text="Conversor de Excel para TXT", font=("Arial", 14, "bold"))
label_titulo.pack(pady=10)

btn_selecionar_arquivo = tk.Button(janela, text="Selecionar Arquivo Excel", command=selecionar_arquivo)
btn_selecionar_arquivo.pack(pady=5)

label_arquivo = tk.Label(janela, text="Nenhum arquivo selecionado.", fg="gray")
label_arquivo.pack(pady=5)

label_colunas = tk.Label(janela, text="Digite os nomes das colunas separados por vírgula:")
label_colunas.pack(pady=5)

entry_colunas = tk.Entry(janela, width=50)
entry_colunas.pack(pady=5)

btn_processar = tk.Button(janela, text="Converter para TXT", command=processar_arquivo, bg="green", fg="white")
btn_processar.pack(pady=10)

# Iniciar o loop da interface
janela.mainloop()
