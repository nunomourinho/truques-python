
# Tutorial de Criação de Documentos do Word com Python

Este tutorial aborda como criar, modificar e extrair informações de documentos do Microsoft Word usando a biblioteca `python-docx` em Python.

## Instalação da Biblioteca

Antes de tudo, você precisa instalar a biblioteca. Use o seguinte comando pip:

```bash
pip install python-docx
```

## Criando um Novo Documento

Importe a biblioteca e use a função `Document` para criar um novo documento:

```python
from docx import Document

# Cria um novo documento
doc = Document()

# Adiciona um título ao documento
doc.add_heading('Título do Documento', level=0)

# Adiciona um parágrafo
p = doc.add_paragraph('Este é um parágrafo simples no documento.')

# Adiciona formatação de texto no mesmo parágrafo
p.add_run(' Este texto será em negrito.').bold = True
p.add_run(' E este será em itálico.').italic = True

# Salva o documento
doc.save('meu_documento.docx')
```

## Adicionando Mais Elementos

### Adicionar Títulos

```python
doc.add_heading('Cabeçalho Nível 1', level=1)
doc.add_heading('Cabeçalho Nível 2', level=2)
```

### Adicionar Listas

```python
doc.add_paragraph('Item 1', style='List Bullet')
doc.add_paragraph('Item 2', style='List Bullet')
doc.add_paragraph('Subitem 2.1', style='List Number')
```

### Adicionar Tabelas

```python
table = doc.add_table(rows=2, cols=2)
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Cabeçalho 1'
hdr_cells[1].text = 'Cabeçalho 2'
row_cells = table.rows[1].cells
row_cells[0].text = 'Conteúdo 1'
row_cells[1].text = 'Conteúdo 2'
```

### Adicionar Imagens

```python
doc.add_picture('caminho_para_imagem.jpg', width=docx.shared.Inches(1.25))
```

## Salvando o Documento

Para salvar o documento:

```python
doc.save('nome_do_documento.docx')
```

# Cheatsheet

## Comandos Básicos

### Instalação

```bash
pip install python-docx
```

### Criar Documento

```python
from docx import Document
doc = Document()
```

### Adicionar Título

```python
doc.add_heading('Título', level=0)
```

### Adicionar Parágrafo

```python
p = doc.add_paragraph('Texto inicial.')
p.add_run(' Texto em negrito.').bold = True
```

### Adicionar Lista

```python
doc.add_paragraph('Item 1', style='List Bullet')
```

### Adicionar Tabela

```python
table = doc.add_table(rows=1, cols=2)
table.rows[0].cells[0].text = 'Célula 1'
```

### Adicionar Imagem

```python
doc.add_picture('caminho_para_imagem.jpg', width=docx.shared.Inches(1))
```

### Salvar Documento

```python
doc.save('meu_documento.docx')
```
