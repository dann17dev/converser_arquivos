import pandas as pd
import os
import time
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from docx import Document
from pptx import Presentation
from tkinter import Tk, Frame, Label, Button, StringVar, Entry, filedialog, OptionMenu, messagebox
from win32com import client  # Para conversão de DOCX e PPTX para PDF

# Função para converter arquivos Excel, CSV e ODS
def convert_to(format_out, input_file, output_file):
    file_ext = os.path.splitext(input_file)[1].lower()

    if file_ext in [".xlsx", ".xlsm"]:
        df = pd.read_excel(input_file, engine='openpyxl')
    elif file_ext == ".xls":
        df = pd.read_excel(input_file, engine='xlrd')
    elif file_ext == ".csv":
        df = pd.read_csv(input_file)
    elif file_ext == ".ods":
        df = pd.read_excel(input_file, engine='odf')
    elif file_ext in [".docx", ".doc"]:
        if format_out == "pdf":
            convert_word_to_pdf(input_file, output_file)
            return True
        text = convert_word_to_text(input_file)
        if format_out == "txt":
            with open(output_file, 'w') as f:
                f.write(text)
            return True
        else:
            print("Conversão de DOCX/DOC para esse formato não suportada.")
            return False
    elif file_ext in [".pptx", ".ppt"]:
        if format_out == "pdf":
            convert_ppt_to_pdf(input_file, output_file)
            return True
        text = convert_ppt_to_text(input_file)
        if format_out == "txt":
            with open(output_file, 'w') as f:
                f.write(text)
            return True
        else:
            print("Conversão de PPTX/PPT para esse formato não suportada.")
            return False
    else:
        print("Formato de entrada não suportado.")
        return False

    # Salvar no formato desejado para DataFrames
    if format_out == "csv":
        df.to_csv(output_file, index=False)
    elif format_out == "xlsx":
        df.to_excel(output_file, index=False)
    elif format_out == "ods":
        df.to_excel(output_file, engine='odf', index=False)
    elif format_out == "pdf":
        convert_to_pdf(df, output_file)
    
    return True

# Função para converter DataFrame em PDF
def convert_to_pdf(df, output_file):
    c = canvas.Canvas(output_file, pagesize=letter)
    width, height = letter
    text = c.beginText(40, height - 40)
    text.setFont("Helvetica", 10)

    # Adicionar dados ao PDF
    for row in df.values:
        text.textLine(','.join(map(str, row)))
    
    c.drawText(text)
    c.showPage()
    c.save()

# Função para converter arquivos Word para texto
def convert_word_to_text(input_file):
    doc = Document(input_file)
    text = []
    for paragraph in doc.paragraphs:
        text.append(paragraph.text)
    return '\n'.join(text)

# Função para converter arquivos PowerPoint para texto
def convert_ppt_to_text(input_file):
    prs = Presentation(input_file)
    text = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text.append(shape.text)
    return '\n'.join(text)

# Função para converter Word para PDF
def convert_word_to_pdf(input_file, output_file):
    word = client.Dispatch('Word.Application')
    doc = word.Documents.Open(input_file)
    doc.SaveAs(output_file, FileFormat=17)  # 17 é o formato PDF
    doc.Close()
    word.Quit()

# Função para converter PowerPoint para PDF
def convert_ppt_to_pdf(input_file, output_file):
    ppt = client.Dispatch('PowerPoint.Application')
    presentation = ppt.Presentations.Open(input_file)
    presentation.SaveAs(output_file, 32)  # 32 é o formato PDF
    presentation.Close()
    ppt.Quit()

# Função para tratar a conversão
def converter(input_file, format_out, output_file):
    if convert_to(format_out, input_file, output_file):
        messagebox.showinfo("Sucesso", f"Arquivo convertido para {output_file} com sucesso.")
        return True
    else:
        messagebox.showerror("Erro", "Conversão falhou.")
        return False

# Função para abrir a caixa de diálogo de seleção de arquivo
def selecionar_arquivo():
    return filedialog.askopenfilename()

# Função para o loop principal de conversão
def iniciar_conversao():
    formato_saida = format_var.get()
    if not formato_saida:
        messagebox.showerror("Erro", "Por favor, selecione um formato de conversão.")
        return

    input_file = selecionar_arquivo()
    if not input_file:
        return

    file_ext = os.path.splitext(input_file)[1].lower()
    supported_formats = [".xlsx", ".xls", ".csv", ".ods", ".docx", ".doc", ".pptx", ".ppt"]

    if file_ext not in supported_formats:
        messagebox.showerror("Erro", f"O arquivo não pode ser convertido. Regras de Conversão:\n"
                                       "- Excel (XLSX, XLS) para CSV, ODS, PDF\n"
                                       "- CSV para XLSX, ODS, PDF\n"
                                       "- ODS para CSV, XLSX, PDF\n"
                                       "- Word (DOCX, DOC) para TXT, PDF\n"
                                       "- PowerPoint (PPTX, PPT) para TXT, PDF\n"
                                       "Não suportamos conversões de outros formatos.")
        return

    time.sleep(2)  # Aguardar 2 segundos

    nome_arquivo = os.path.basename(input_file)
    nome_novo_arquivo = f"novo arquivo - {nome_arquivo}"
    entry_nome_arquivo.delete(0, 'end')
    entry_nome_arquivo.insert(0, nome_novo_arquivo)

    local_salvar = filedialog.askdirectory(title="Escolha o local para salvar o arquivo", initialdir=os.path.expanduser("~/Downloads"))

    output_file = os.path.join(local_salvar, entry_nome_arquivo.get())
    
    output_file += f".{formato_saida}"

    confirm_message = f"O arquivo {input_file} será convertido para o formato {formato_saida}. Deseja continuar?"
    if messagebox.askyesno("Confirmação", confirm_message):
        converter(input_file, formato_saida, output_file)

# Configurações da interface gráfica
def main():
    global entry_nome_arquivo, format_var

    root = Tk()
    root.title("Conversor de Arquivos")
    root.geometry("400x600")
    root.configure(bg="#F5F5F5")  # Fundo cinza claro

    # Mensagem de boas-vindas
    welcome_label = Label(root, text="Seja bem-vindo ao conversor!", font=("Helvetica", 14), bg="#F5F5F5", fg="#333333")
    welcome_label.pack(pady=10)

    # Tabela de possibilidades de conversão
    possibilities_frame = Frame(root, bg="#F5F5F5")
    possibilities_frame.pack(pady=10)

    Label(possibilities_frame, text="Possibilidades de Conversão:", font=("Helvetica", 12, 'bold'), bg="#F5F5F5", fg="#333333").grid(row=0, columnspan=2, pady=5)

    formats = [("Excel (XLSX, XLS)", "CSV, ODS, PDF"),
               ("CSV", "XLSX, ODS, PDF"),
               ("ODS", "CSV, XLSX, PDF"),
               ("Word (DOCX, DOC)", "TXT, PDF"),
               ("PowerPoint (PPTX, PPT)", "TXT, PDF")]

    for idx, (input_format, output_format) in enumerate(formats):
        Label(possibilities_frame, text=input_format, font=("Helvetica", 10), bg="#F5F5F5", fg="#333333").grid(row=idx+1, column=0, padx=10, pady=5, sticky="w")
        Label(possibilities_frame, text=output_format, font=("Helvetica", 10), bg="#F5F5F5", fg="#333333").grid(row=idx+1, column=1, padx=10, pady=5, sticky="w")

    # Texto de regras de conversão
    rules_text = "Regras de Conversão:\n- Excel (XLSX, XLS) para CSV, ODS, PDF\n" \
                 "- CSV para XLSX, ODS, PDF\n" \
                 "- ODS para CSV, XLSX, PDF\n" \
                 "- Word (DOCX, DOC) para TXT, PDF\n" \
                 "- PowerPoint (PPTX, PPT) para TXT, PDF\n\n" \
                 "Qualquer outra conversão não é suportada."

    rules_info = Label(root, text=rules_text, font=("Helvetica", 10), bg="#F5F5F5", fg="#333333", wraplength=380, justify="left")
    rules_info.pack(pady=10)

    # Seleção de formato
    format_var = StringVar(root)
    format_var.set("csv")

    formats = ["csv", "xlsx", "ods", "pdf", "txt"]
    format_menu = OptionMenu(root, format_var, *formats)
    format_menu.config(bg="#4A90E2", fg="white", font=("Helvetica", 12), width=20)
    format_menu.pack(pady=10)

    # Campo para nome do arquivo
    entry_nome_arquivo = Entry(root, font=("Helvetica", 12), width=30, bg="#FFFFFF", fg="#333333", bd=2, relief="groove")
    entry_nome_arquivo.pack(pady=10)

    # Botão de conversão
    convert_button = Button(root, text="Escolher arquivo para conversão", command=iniciar_conversao, bg="#4A90E2", fg="white", font=("Helvetica", 12), width=30)
    convert_button.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
