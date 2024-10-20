import os
import time
from tkinter import Tk, Label, Button, StringVar, OptionMenu, Entry, filedialog, messagebox
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from docx import Document
from pptx import Presentation
from win32com import client


def convert_to(format_out, input_file, output_file):
    """
    Converte arquivos de diferentes formatos (Excel, CSV, ODS, DOCX, PPTX)
    para o formato desejado (CSV, ODS, PDF, TXT).
    """
    file_ext = os.path.splitext(input_file)[1].lower()

    # Carrega o DataFrame conforme a extensão do arquivo de entrada
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

    # Salva no formato desejado para DataFrames
    if format_out == "csv":
        df.to_csv(output_file, index=False)
    elif format_out == "xlsx":
        df.to_excel(output_file, index=False)
    elif format_out == "ods":
        df.to_excel(output_file, engine='odf', index=False)
    elif format_out == "pdf":
        convert_to_pdf(df, output_file)

    return True


def convert_to_pdf(df, output_file):
    """Converte um DataFrame para um arquivo PDF."""
    c = canvas.Canvas(output_file, pagesize=letter)
    width, height = letter
    text = c.beginText(40, height - 40)
    text.setFont("Helvetica", 10)

    # Adiciona dados ao PDF
    for row in df.values:
        text.textLine(','.join(map(str, row)))

    c.drawText(text)
    c.showPage()
    c.save()


def convert_word_to_text(input_file):
    """Converte um arquivo Word para texto."""
    doc = Document(input_file)
    text = [paragraph.text for paragraph in doc.paragraphs]
    return '\n'.join(text)


def convert_ppt_to_text(input_file):
    """Converte um arquivo PowerPoint para texto."""
    prs = Presentation(input_file)
    text = [shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")]
    return '\n'.join(text)


def convert_word_to_pdf(input_file, output_file):
    """Converte um arquivo Word para PDF usando a biblioteca win32com."""
    word = client.Dispatch('Word.Application')
    doc = word.Documents.Open(input_file)
    doc.SaveAs(output_file, FileFormat=17)  # 17 é o formato PDF
    doc.Close()
    word.Quit()


def convert_ppt_to_pdf(input_file, output_file):
    """Converte um arquivo PowerPoint para PDF usando a biblioteca win32com."""
    ppt = client.Dispatch('PowerPoint.Application')
    presentation = ppt.Presentations.Open(input_file)
    presentation.SaveAs(output_file, 32)  # 32 é o formato PDF
    presentation.Close()
    ppt.Quit()


def converter(input_file, format_out, output_file):
    """Executa a conversão e exibe uma mensagem de sucesso ou erro."""
    if convert_to(format_out, input_file, output_file):
        messagebox.showinfo("Sucesso", f"Arquivo convertido para {output_file} com sucesso.")
        return True
    else:
        messagebox.showerror("Erro", "Conversão falhou.")
        return False


def selecionar_arquivo():
    """Abre a caixa de diálogo para selecionar um arquivo."""
    return filedialog.askopenfilename()


def iniciar_conversao():
    """Inicia o processo de conversão ao clicar no botão."""
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
                                       "Qualquer outra conversão não é suportada.*")
        return

    time.sleep(2)  # Aguardar 2 segundos

    nome_arquivo = os.path.basename(input_file)
    nome_novo_arquivo = f"novo_arquivo - {nome_arquivo}"
    entry_nome_arquivo.delete(0, 'end')
    entry_nome_arquivo.insert(0, nome_novo_arquivo)

    local_salvar = filedialog.askdirectory(title="Escolha o local para salvar o arquivo",
                                            initialdir=os.path.expanduser("~/Downloads"))

    output_file = os.path.join(local_salvar, entry_nome_arquivo.get())
    output_file += f".{formato_saida}"

    confirm_message = f"O arquivo {input_file} será convertido para o formato {formato_saida}. Deseja continuar?"
    if messagebox.askyesno("Confirmação", confirm_message):
        converter(input_file, formato_saida, output_file)


def main():
    """Configura e inicia a interface gráfica."""
    global entry_nome_arquivo, format_var

    root = Tk()
    root.title("Conversor de Arquivos")
    root.geometry("400x600")
    root.configure(bg="#F5F5F5")  # Fundo cinza claro

    # Mensagem de boas-vindas
    welcome_label = Label(root, text="Seja bem-vindo ao conversor!", font=("Helvetica", 13, 'bold'),
                          bg="#F5F5F5", fg="#333333")
    welcome_label.pack(pady=10)

    # Apresentação do sistema
    presentation_label = Label(root, text="É um prazer ter você acessando nosso sistema de conversão! "
                                            "Ele te ajudará a fazer em minutos algo que demoraria muito e "
                                            "deixaria você menos produtivo.", font=("Helvetica", 12),
                                            bg="#F5F5F5", fg="#333333", wraplength=380, justify="left")
    presentation_label.pack(pady=10)

    # Seleção de formato
    format_label = Label(root, text="Selecione o formato do novo arquivo:", font=("Helvetica", 12, 'bold'),
                         bg="#F5F5F5", fg="#333333")
    format_label.pack(pady=10)

    format_var = StringVar(root)
    format_var.set("csv")

    formats = ["csv", "xlsx", "ods", "pdf", "txt"]
    format_menu = OptionMenu(root, format_var, *formats)
    format_menu.config(bg="#007ACC", fg="white", font=("Ubuntu Mono", 12), width=20)
    format_menu.pack(pady=10)

    # Campo para nome do arquivo
    entry_nome_arquivo = Entry(root, font=("Ubuntu Mono", 12), width=30, bg="#FFFFFF",
                                fg="#333333", bd=2, relief="groove")
    entry_nome_arquivo.insert(0, "novo_arquivo")
    entry_nome_arquivo.pack(pady=10)

    # Botão de conversão
    convert_button = Button(root, text="Iniciar Conversão", command=iniciar_conversao,
                            bg="#007ACC", fg="white", font=("Ubuntu Mono", 12),
                            width=30, borderwidth=2, relief="groove")
    convert_button.pack(pady=20)

    # Regras de conversão
    rules_label = Label(root, text="Regras de Conversão:", font=("Helvetica", 13, 'bold'),
                        bg="#F5F5F5", fg="#333333")
    rules_label.pack(pady=10)

    rules_text = Label(root, text="- Excel (XLSX, XLS) para CSV, ODS, PDF\n"
                                   "- CSV para XLSX, ODS, PDF\n"
                                   "- ODS para CSV, XLSX, PDF\n"
                                   "- Word (DOCX, DOC) para TXT, PDF\n"
                                   "- PowerPoint (PPTX, PPT) para TXT, PDF\n"
                                   "Qualquer outra conversão não é suportada.", font=("Helvetica", 11),
                       bg="#F5F5F5", fg="#333333", wraplength=380, justify="left")
    rules_text.pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    main()
