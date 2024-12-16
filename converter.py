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
        carregar_colunas()
    else:
        label_arquivo['text'] = "Nenhum arquivo selecionado."

def carregar_colunas():
    """Carrega as colunas do arquivo Excel e exibe na Listbox"""
    try:
        df = pd.read_excel(excel_file, engine='openpyxl')
        print("Colunas carregadas:", df.columns.tolist())  # Debugging line
        listbox_colunas.delete(0, tk.END)  # Limpa o Listbox
        for coluna in df.columns:
            listbox_colunas.insert(tk.END, coluna)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar colunas: {str(e)}")

# Variável global para armazenar as colunas selecionadas
colunas_selecionadas = []

def adicionar_coluna():
    """Adiciona a coluna selecionada à lista de colunas selecionadas"""
    global colunas_selecionadas
    selecionado = listbox_colunas.curselection()
    if selecionado:
        coluna = listbox_colunas.get(selecionado)
        if coluna not in colunas_selecionadas:
            colunas_selecionadas.append(coluna)
            atualizar_listbox_adicionadas()

def remover_coluna():
    """Remove a coluna selecionada da lista de colunas selecionadas"""
    global colunas_selecionadas
    selecionado = listbox_colunas_adicionadas.curselection()
    if selecionado:
        coluna = listbox_colunas_adicionadas.get(selecionado)
        colunas_selecionadas.remove(coluna)
        atualizar_listbox_adicionadas()

def atualizar_listbox_adicionadas():
    """Atualiza a Listbox de colunas adicionadas"""
    listbox_colunas_adicionadas.delete(0, tk.END)
    for coluna in colunas_selecionadas:
        listbox_colunas_adicionadas.insert(tk.END, coluna)

def formatar_dados(df):
    """
    Formata os campos 'MES_ANO_DIREITO' (data), 'CPF' e 'CNPJ' no DataFrame.
    """
    # Formatar a coluna de data
    if 'MES_ANO_DIREITO' in df.columns:
        df['MES_ANO_DIREITO'] = pd.to_datetime(df['MES_ANO_DIREITO'], errors='coerce')
        df['MES_ANO_DIREITO'] = df['MES_ANO_DIREITO'].fillna("").dt.strftime('%d/%m/%Y')

    # Formatar o CPF
    if 'CPF' in df.columns:
        df['CPF'] = df['CPF'].fillna("").astype(str).str.replace(r'[^\d]', '', regex=True)  # Remove caracteres não numéricos
        
        # Garante que o CPF tenha exatamente 11 dígitos
        df['CPF'] = df['CPF'].apply(lambda x: x[:11] if len(x) > 11 else x.zfill(11))  # Trunca ou preenche com zeros
        
        # Aplica a máscara padrão
        df['CPF'] = df['CPF'].str.replace(
            r'(\d{3})(\d{3})(\d{3})(\d{2})', r'\1.\2.\3-\4', regex=True
        )  #  # Aplica a máscara padrão

    # Formatar o CNPJ
    if 'CNPJ' in df.columns:
        df['CNPJ'] = df['CNPJ'].fillna("").astype(str).str.replace(r'[^\d]', '', regex=True)
        df['CNPJ'] = df['CNPJ'].apply(lambda x: x.zfill(14) if len(x) < 14 else x)  # Garante 14 dígitos
        df['CNPJ'] = df['CNPJ'].str.replace(
            r'(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})', r'\1.\2.\3/\4-\5', regex=True
        )  # Aplica a máscara padrão

    # Converter colunas numéricas para int, se possível
    for col in df.select_dtypes(include=['float']):
        df[col] = df[col].fillna(0).astype(int)  # Converte para int, substituindo NaN por 0

    return df

def processar_arquivo():
    """Processa o arquivo Excel e salva como TXT"""
    try:
        if not excel_file:
            raise ValueError("Por favor, selecione um arquivo Excel.")

        df = pd.read_excel(excel_file, engine='openpyxl')

        if not colunas_selecionadas:
            raise ValueError("Por favor, selecione pelo menos uma coluna.")

        # Verifica se todas as colunas selecionadas estão no DataFrame
        colunas_invalidas = [col for col in colunas_selecionadas if col not in df.columns]
        if colunas_invalidas:
            raise ValueError(f"As seguintes colunas não estão no arquivo: {', '.join(colunas_invalidas)}")

        # Filtrar as colunas na ordem selecionada
        df_filtrado = df[colunas_selecionadas]

        df_formatado = df_filtrado.drop_duplicates()
        # Formatar os dados
        df_formatado = formatar_dados(df_formatado)

        output_file = filedialog.asksaveasfilename(
            title="Salvar arquivo como",
            defaultextension=".txt",
            filetypes=[("Arquivo de Texto", "*.txt")]
        )
        if not output_file:
            raise ValueError("Nenhum caminho para salvar foi selecion.")

        df_formatado.to_csv(output_file, sep=';', index=False)

        messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em:\n{output_file}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

# Criar a janela principal
janela = tk.Tk()
janela.title("Conversor de Excel para TXT")

excel_file = None

label_titulo = tk.Label(janela, text="Conversor de Excel para TXT", font=("Arial", 14, "bold"))
label_titulo.pack(pady=10)

btn_selecionar_arquivo = tk.Button(janela, text="Selecionar Arquivo Excel", command=selecionar_arquivo)
btn_selecionar_arquivo.pack(pady=5)

label_arquivo = tk.Label(janela, text="Nenhum arquivo selecionado.", fg="gray")
label_arquivo.pack(pady=5)

label_colunas = tk.Label(janela, text="Selecione as colunas desejadas:")
label_colunas.pack(pady=5)

listbox_colunas = tk.Listbox(janela, width=50, height=10)
listbox_colunas.pack(pady=5)

btn_adicionar = tk.Button(janela, text="Adicionar", command=adicionar_coluna)
btn_adicionar.pack(pady=5)

label_colunas_adicionadas = tk.Label(janela, text="Colunas adicionadas:")
label_colunas_adicionadas.pack(pady=5)

listbox_colunas_adicionadas = tk.Listbox(janela, width=50, height=10)
listbox_colunas_adicionadas.pack(pady=5)

btn_remover = tk.Button(janela, text="Remover", command=remover_coluna)
btn_remover.pack(pady=5)

btn_processar = tk.Button(janela, text="Converter para TXT", command=processar_arquivo, bg="green", fg="white")
btn_processar.pack(pady=10)

# Iniciar o loop da interface
janela.mainloop()
