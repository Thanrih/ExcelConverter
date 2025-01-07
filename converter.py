import pandas as pd
import os
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

    for coluna in df.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]']).columns:
            df[coluna] = df[coluna].apply(lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "")

    if 'CPF' in df.columns:
    # Substitui valores NaN por uma string vazia e remove caracteres não numéricos
        df['CPF'] = df['CPF'].fillna("").astype(str).str.replace(r'[^\d]', '', regex=True)
        
        # Remove o último caractere, se existir
        df['CPF'] = df['CPF'].apply(lambda x: x[:-1] if len(x) > 0 else x)
        
        # Garante que o CPF tenha exatamente 11 dígitos (trunca ou adiciona zeros à esquerda)
        def ajustar_cpf(cpf):
            cpf = cpf[:11] if len(cpf) > 11 else ('0' * (11 - len(cpf))) + cpf
            return cpf
        
        df['CPF'] = df['CPF'].apply(ajustar_cpf)
        print(df['CPF'])
        # Aplica a máscara padrão do CPF
        def aplicar_mascara(cpf):
            return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}" if len(cpf) == 11 else cpf
        
        df['CPF'] = df['CPF'].apply(aplicar_mascara)



    if 'CNPJ' in df.columns:
        # Remove caracteres não numéricos
        df.loc[:, 'CNPJ'] = df['CNPJ'].fillna("").astype(str).str.replace(r'[^\d]', '', regex=True)
        # Remove CNPJs inválidos (com mais de 14 dígitos ou vazios)
        df.loc[:, 'CNPJ'] = df['CNPJ'].apply(lambda x: x if len(x) <= 14 else x[:14])
        # Adiciona zeros à esquerda para garantir 14 dígitos
        df.loc[:, 'CNPJ'] = df['CNPJ'].apply(lambda x: x.zfill(14))
        # Aplica a máscara de CNPJ
        df.loc[:, 'CNPJ'] = df['CNPJ'].str.replace(r'(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})', r'\1.\2.\3/\4-\5', regex=True)

    # Remove duplicatas novamente após formatações

    return df




def processar_arquivo():
    """Processa o arquivo Excel e salva como TXT com ; no final de cada linha"""
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

        # Formatar os dados
        df_formatado = formatar_dados(df_filtrado)

        # Adiciona ';' ao final de cada linha, ignorando linhas idênticas
        output_file = filedialog.asksaveasfilename(
            title="Salvar arquivo como",
            defaultextension=".txt",
            filetypes=[("Arquivo de Texto", "*.txt")]
        )
        if not output_file:
            raise ValueError("Nenhum caminho para salvar foi selecion.")

        # Salva o arquivo com separador ';' e preserva caracteres especiais
        temp_file = output_file + "_temp"
        df_formatado.to_csv(temp_file, sep=';', index=False, header=True)
        with open(temp_file, 'r', encoding='utf-8') as infile, open(output_file, 'w', encoding='utf-8') as outfile:
            previous_line = None
            for i, line in enumerate(infile):
                if i == 0:  # Primeira linha (cabeçalho)
                    outfile.write(line.strip() + ';\n')
                    print(i)
                else:  # Demais linhas (dados)
                    if line.strip() != previous_line:  # Ignora linhas idênticas
                        outfile.write(line.strip() + ';\n')
                    previous_line = line.strip()

        messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em:\n{output_file}")
        os.remove(temp_file)

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
