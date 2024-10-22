# Conversor de Arquivos

Este projeto é um **Conversor de Arquivos** que permite a conversão de diversos formatos, como Excel, CSV, Word e PowerPoint, para outros formatos populares. O objetivo principal é fornecer uma ferramenta fácil de usar que possibilite a conversão de arquivos de forma eficiente, atendendo a diferentes necessidades.

## Funcionalidades

### Conversão de formatos suportados:

- **Excel (XLSX, XLS)** para CSV, ODS, PDF
- **CSV** para XLSX, ODS, PDF
- **ODS** para CSV, XLSX, PDF
- **Word (DOCX, DOC)** para TXT, PDF
- **PowerPoint (PPTX, PPT)** para PDF

### Interface Gráfica:
O aplicativo possui uma interface gráfica amigável, permitindo que os usuários selecionem arquivos, escolham o formato de saída, nomeiem o novo arquivo e escolham o local de salvamento.

### Notificações:
Mensagens informativas que notificam o usuário sobre o progresso da conversão e possíveis erros.

## Requisitos

Para executar este projeto, você precisará das seguintes bibliotecas Python instaladas:

- `pandas`: Manipulação de dados e leitura/escrita de arquivos.
- `docx`: Manipulação de arquivos do Microsoft Word.
- `ttk`: Criação da interface gráfica com barra de progresso.
- `tkinter`: Criação da interface gráfica.
- `openpyxl`: Manipulação de arquivos Excel.
- `odf`: Manipulação de arquivos OpenDocument (ODS).
- `wkhtmltopdf`: Conversão de arquivos HTML para PDF.
- `libreoffice`: Conversão de documentos para PDF (somente para arquivos DOC, DOCX, PPT e PPTX).

### Como funciona:

- A primeira seção do arquivo (`# Conversor de Arquivos`) descreve o objetivo do projeto.
- A seção "Funcionalidades" detalha as possíveis conversões e a interface gráfica.
- Em "Requisitos", são listadas as bibliotecas necessárias e um comando para instalação.
- A seção "Como Usar" fornece um guia passo a passo para utilização do programa.
- O link para download do instalador é incluído ao final.

Isso dá uma visão clara do projeto, facilitando para outros usuários entenderem e baixarem o instalador.
[Download do Instalador](https://drive.google.com/file/d/1-Altw0k8PnuiYmttNZ1mvn65CkIIaqXu/view?usp=sharing)

Você pode instalar as dependências necessárias usando o seguinte comando:

```bash
pip install pandas python-docx openpyxl odf python-tk
