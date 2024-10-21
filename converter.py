import os
import pandas as pd
from tkinter import Tk, Label, Button, Entry, filedialog, StringVar, OptionMenu, messagebox, ttk

# Função para selecionar arquivo
def selecionar_arquivo():
    global input_file
    input_file = filedialog.askopenfilename(title="Selecione um arquivo", filetypes=(("Todos os arquivos", "*.*"),))
    if input_file:
        nome_arquivo = os.path.basename(input_file)
        nome_novo_arquivo = f"novo_arquivo - {nome_arquivo}"
        entry_nome_arquivo.delete(0, 'end')
        entry_nome_arquivo.insert(0, nome_novo_arquivo)

# Função para ajustar o DataFrame
def ajustar_dataframe(df):
    # Exemplo de ajuste: renomear colunas
    df.rename(columns={'Nome Antigo': 'Nome Novo'}, inplace=True)
    
    # Exemplo de filtragem: remover linhas com valores ausentes em uma coluna específica
    df.dropna(subset=['Nome Novo'], inplace=True)

    # Exemplo de adição de nova coluna
    if 'Idade' in df.columns:
        df['Idade em 10 Anos'] = df['Idade'] + 10

    return df

# Função para converter o arquivo
def converter(input_file, formato_saida, output_file):
    try:
        # Lógica de conversão para diferentes formatos
        if input_file.endswith(('.xlsx', '.xls')) and formato_saida == "csv":
            df = pd.read_excel(input_file, engine='openpyxl')
            df = ajustar_dataframe(df)  # Ajustar DataFrame
            df.to_csv(output_file, index=False)
        elif input_file.endswith('.csv') and formato_saida == "xlsx":
            df = pd.read_csv(input_file)
            df = ajustar_dataframe(df)  # Ajustar DataFrame
            df.to_excel(output_file, index=False, engine='openpyxl')
        elif input_file.endswith(('.ods', '.xlsx', '.xls')) and formato_saida == "pdf":
            df = pd.read_excel(input_file, engine='openpyxl')
            df = ajustar_dataframe(df)  # Ajustar DataFrame
            df.to_html('temp.html')
            os.system(f"wkhtmltopdf temp.html {output_file}")
            os.remove('temp.html')
        else:
            messagebox.showerror("Erro", "Formato de conversão não suportado.")
            return False
        
        messagebox.showinfo("Sucesso", f"Arquivo convertido com sucesso: {output_file}")
        return True
    except Exception as e:
        messagebox.showerror("Erro", f"Falha na conversão: {str(e)}")
        return False

# Função que inicia a conversão após confirmação
def iniciar_conversao():
    formato_saida = formato_var.get()
    if not formato_saida:
        messagebox.showerror("Erro", "Selecione um formato de saída.")
        return
    local_salvar = filedialog.askdirectory(title="Escolha o local para salvar o arquivo",
                                           initialdir=os.path.expanduser("~/Downloads"))
    if not local_salvar:
        local_salvar = os.path.expanduser("~/Downloads")  # Diretório padrão
    
    output_file = os.path.join(local_salvar, entry_nome_arquivo.get())
    output_file += f".{formato_saida}"
    
    if converter(input_file, formato_saida, output_file):
        os.startfile(local_salvar)  # Abre a pasta onde o arquivo foi salvo

# Interface Gráfica (UI)
root = Tk()
root.title("Conversor de Arquivos")
root.geometry("500x600")
root.config(bg="#F5F5F5")

# Mensagem de boas-vindas
welcome_label = Label(root, text="Seja bem-vindo ao conversor!", font=("Helvetica", 13, 'bold'),
                      bg="#F5F5F5", fg="#333333")
welcome_label.pack(pady=10)

# Resumo das opções de conversão
summary_label = Label(root, text="Resumo de conversões possíveis:\n"
                                 "- Excel (XLSX, XLS) para CSV, ODS, PDF\n"
                                 "- CSV para XLSX, ODS, PDF\n"
                                 "- ODS para CSV, XLSX, PDF\n"
                                 "- Word (DOCX, DOC) para TXT, PDF\n"
                                 "- PowerPoint (PPTX, PPT) para TXT, PDF\n",
                      font=("Helvetica", 11), bg="#F5F5F5", fg="#333333", justify="left")
summary_label.pack(pady=5)

# Botão para selecionar o arquivo
selecionar_arquivo_button = Button(root, text="Selecionar Arquivo", command=selecionar_arquivo, width=30, height=2,
                                   bg="#191970", fg="white", font=("Montserrat", 12, 'bold'))
selecionar_arquivo_button.pack(pady=10)

# Campo para inserir o nome do arquivo de saída
entry_nome_arquivo = Entry(root, font=("Montserrat", 12), width=30, bg="#FFFFFF", fg="#333333", bd=2, relief="groove")
entry_nome_arquivo.insert(0, "novo_arquivo")
entry_nome_arquivo.pack(pady=10)

# Seletor de formato de saída
formatos = ["csv", "xlsx", "pdf", "ods"]
formato_var = StringVar(root)
formato_var.set("")  # Lista vazia por padrão
formato_label = Label(root, text="Escolha o formato de saída", font=("Helvetica", 12), bg="#F5F5F5", fg="#333333")
formato_label.pack(pady=5)
formato_menu = OptionMenu(root, formato_var, *formatos)
formato_menu.config(width=30, font=("Montserrat", 12), bg="#F5F5F5", fg="#333333")
formato_menu.pack(pady=10)

# Botão para converter o arquivo
converter_button = Button(root, text="Converter", command=iniciar_conversao, width=30, height=2,
                          bg="#00FF00", fg="black", font=("Montserrat", 12, 'bold'))  # Cor Lime
converter_button.pack(pady=20)

root.mainloop()
