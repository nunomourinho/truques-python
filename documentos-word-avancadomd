
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

## Funcionalidades Avançadas

### Configurar Estilos de Parágrafo

```python
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

doc = Document()
p = doc.add_paragraph('Este é um parágrafo centralizado.')
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(12)
p.paragraph_format.left_indent = Pt(24)
doc.save('paragrafo_estilizado.docx')
```

### Trabalhar com Tabelas

```python
from docx import Document

doc = Document()
table = doc.add_table(rows=3, cols=3)
table.style = 'Table Grid'

# Mesclar células
a = table.cell(0, 0)
b = table.cell(1, 0)
a.merge(b)

# Adicionar conteúdo às células
table.cell(0, 0).text = 'Célula mesclada'
table.cell(0, 1).text = 'Célula 2'
table.cell(0, 2).text = 'Célula 3'
doc.save('tabela.docx')
```

### Adicionar Cabeçalhos e Rodapés

```python
from docx import Document

doc = Document()
section = doc.sections[0]
header = section.header
paragraph = header.paragraphs[0]
paragraph.text = "Cabeçalho do Documento"

footer = section.footer
paragraph = footer.paragraphs[0]
paragraph.text = "Rodapé do Documento"
doc.save('cabecalho_rodape.docx')
```

### Configurar Seções

```python
from docx import Document
from docx.enum.section import WD_ORIENT

doc = Document()
section = doc.sections[-1]
section.orientation = WD_ORIENT.LANDSCAPE
section.page_width = docx.shared.Inches(11)
section.page_height = docx.shared.Inches(8.5)
doc.save('pagina_paisagem.docx')
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

## Funcionalidades Avançadas

### Configurar Parágrafo

```python
from docx.shared import Pt
p = doc.add_paragraph('Texto.')
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(12)
p.paragraph_format.left_indent = Pt(24)
```

### Mesclar Células na Tabela

```python
a = table.cell(0, 0)
b = table.cell(1, 0)
a.merge(b)
```

### Adicionar Cabeçalho e Rodapé

```python
header = section.header
header.paragraphs[0].text = "Cabeçalho"
footer = section.footer
footer.paragraphs[0].text = "Rodapé"
```

### Orientação da Página

```python
section.orientation = WD_ORIENT.LANDSCAPE
section.page_width = docx.shared.Inches(11)
section.page_height = docx.shared.Inches(8.5)
```
