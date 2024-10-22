# Conversor de Arquivos

Este projeto é um conversor de arquivos que permite a conversão entre diversos formatos populares, como Excel, CSV, Word, PowerPoint e outros. O objetivo é fornecer uma ferramenta fácil de usar, com uma interface gráfica amigável, que possibilite a conversão de arquivos de maneira rápida e eficiente.

## Funcionalidades

### Conversões Suportadas:
- **Excel (XLSX, XLS)** para CSV, ODS, PDF
- **CSV** para XLSX, ODS, PDF
- **ODS** para CSV, XLSX, PDF
- **Word (DOCX, DOC)** para TXT, PDF
- **PowerPoint (PPTX, PPT)** para PDF

### Detalhes:
- **Interface Gráfica (Tkinter)**: Simples e intuitiva, permitindo ao usuário:
  - Selecionar um arquivo
  - Escolher o formato de saída
  - Nomear o novo arquivo
  - Definir o local de salvamento

- **Barra de Progresso**: Exibe o andamento da conversão.
- **Mensagens de Alerta**: Notificações que informam o sucesso ou falha na conversão.

## Requisitos

Para executar este projeto, você precisará das seguintes bibliotecas Python instaladas:

- `pandas`: Manipulação de dados e leitura/escrita de arquivos.
- `python-docx`: Manipulação de arquivos do Microsoft Word.
- `tkinter`: Criação da interface gráfica.
- `docx`: Manipulação de arquivos do Microsoft Word.
- `ttk`: Para a barra de progresso e outros widgets.

### Instalação das dependências:

Use o seguinte comando para instalar as dependências:

```bash
pip install pandas python-docx tk


