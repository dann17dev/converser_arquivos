import os
import pandas as pd
from tkinter import Tk, Label, Button, Entry, filedialog, StringVar, OptionMenu, messagebox
from tkinter import ttk
from docx import Document
import threading
import time

# Função para selecionar arquivo
def selecionar_arquivo():
    global input_file
    input_file = filedialog.askopenfilename(title="Selecione um arquivo", filetypes=(("Todos os arquivos", "*.*"),))
    if input_file:
        nome_arquivo = os.path.basename(input_file)
        nome_novo_arquivo = f"novo_arquivo - {nome_arquivo}"
        entry_nome_arquivo.delete(0, 'end')
        entry_nome_arquivo.insert(0, nome_novo_arquivo)

# Função para converter o arquivo
def converter(input_file, formato_saida, output_file, progress_var):
    try:
        # Simulação de passos para a barra de progresso
        steps = 5
        for step in range(steps):
            time.sleep(0.5)  # Simulação de tempo de processamento
            progress_var.set((step + 1) / steps * 100)  # Atualiza a barra de progresso

        # Lógica de conversão para diferentes formatos
        if input_file.endswith(('.xlsx', '.xls')) and formato_saida == "csv":
            df = pd.read_excel(input_file)
            df.to_csv(output_file, index=False)
        elif input_file.endswith('.csv') and formato_saida in ["xlsx", "ods"]:
            df = pd.read_csv(input_file)
            if formato_saida == "xlsx":
                df.to_excel(output_file, index=False)
            elif formato_saida == "ods":
                df.to_excel(output_file, index=False, engine='odf')
        elif input_file.endswith('.ods') and formato_saida in ["csv", "xlsx"]:
            df = pd.read_excel(input_file, engine='odf')
            if formato_saida == "csv":
                df.to_csv(output_file, index=False)
            elif formato_saida == "xlsx":
                df.to_excel(output_file, index=False)
        elif input_file.endswith(('.xlsx', '.xls', '.ods')) and formato_saida == "pdf":
            df = pd.read_excel(input_file, engine='openpyxl' if input_file.endswith('.xlsx') else 'odf')
            df.to_html('temp.html')
            os.system(f"wkhtmltopdf temp.html {output_file}")
            os.remove('temp.html')
        elif input_file.endswith(('.docx', '.doc')) and formato_saida == "txt":
            doc = Document(input_file)
            with open(output_file, 'w', encoding='utf-8') as f:
                for para in doc.paragraphs:
                    f.write(para.text + '\n')
        elif input_file.endswith(('.docx', '.doc')) and formato_saida == "pdf":
            os.system(f"libreoffice --headless --convert-to pdf {input_file} --outdir {os.path.dirname(output_file)}")
        elif input_file.endswith(('.pptx', '.ppt')) and formato_saida == "pdf":
            os.system(f"libreoffice --headless --convert-to pdf {input_file} --outdir {os.path.dirname(output_file)}")
        else:
            messagebox.showerror("Erro", f"Formato de conversão não suportado: {formato_saida} de {input_file}.")
            return False

        messagebox.showinfo("Sucesso", f"Arquivo convertido com sucesso: {output_file}")
        progress_var.set(100)  # Atualiza a barra de progresso para 100%
        return True
    except Exception as e:
        messagebox.showerror("Erro", f"Falha na conversão: {str(e)}")
        progress_var.set(0)  # Reseta a barra de progresso
        return False

# Função que inicia a conversão
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

    # Configura a barra de progresso
    progress_var = StringVar()
    progress_bar = ttk.Progressbar(root, length=300, mode='determinate', variable=progress_var)
    progress_bar.pack(pady=20)

    # Executar a conversão em uma thread para não bloquear a interface
    thread = threading.Thread(target=converter, args=(input_file, formato_saida, output_file, progress_var))
    thread.start()

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
                                 "- PowerPoint (PPTX, PPT) para PDF\n",
                      font=("Helvetica", 11), bg="#F5F5F5", fg="#333333", justify="left")
summary_label.pack(pady=5)

# Botão de seleção de arquivo
selecionar_arquivo_button = Button(root, text="Selecionar Arquivo", command=selecionar_arquivo, width=30, height=2,
                                   bg="#191970", fg="white", font=("Montserrat", 12, 'bold'))
selecionar_arquivo_button.pack(pady=10)

# Campo para inserir o nome do arquivo de saída
entry_nome_arquivo = Entry(root, font=("Montserrat", 12), width=30, bg="#FFFFFF", fg="#333333", bd=2, relief="groove")
entry_nome_arquivo.insert(0, "novo_arquivo")
entry_nome_arquivo.pack(pady=10)

# Seletor de formato de saída
formatos = ["csv", "xlsx", "ods", "pdf", "txt"]
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
