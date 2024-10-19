# Conversor de Arquivos

Este projeto é um conversor de arquivos que permite a conversão de diversos formatos, como Excel, CSV, Word e PowerPoint, para outros formatos populares. O objetivo principal é fornecer uma ferramenta fácil de usar que possibilite a conversão de arquivos de forma eficiente, atendendo a diferentes necessidades.

## Funcionalidades

- **Conversão de formatos suportados**:
  - Excel (XLSX, XLSM, XLS, CSV) para CSV, XLSX, ODS, PDF
  - CSV para CSV, XLSX, ODS, PDF
  - Word (DOCX, DOC) para TXT, PDF
  - PowerPoint (PPTX, PPT) para TXT, PDF

- **Interface Gráfica**: O aplicativo possui uma interface gráfica amigável, permitindo que os usuários selecionem arquivos, escolham o formato de saída, nomeiem o novo arquivo e escolham o local de salvamento.

- **Notificações**: Mensagens informativas que notificam o usuário sobre o progresso da conversão e possíveis erros.

## Requisitos

Para executar este projeto, você precisará das seguintes bibliotecas Python instaladas:

- `pandas`: Manipulação de dados e leitura/escrita de arquivos.
- `reportlab`: Geração de arquivos PDF.
- `python-docx`: Manipulação de arquivos do Microsoft Word.
- `python-pptx`: Manipulação de arquivos do Microsoft PowerPoint.
- `pywin32`: Acesso às APIs do Windows.
- `tkinter`: Criação da interface gráfica.

Você pode instalar as dependências necessárias usando o seguinte comando:

```bash
pip install pandas reportlab python-docx python-pptx pywin32
