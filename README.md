# python-docx

*python-docx* is a Python library for reading, creating, and updating Microsoft Word 2007+ (.docx) files.

This repository exists as a fork of [the official repo](https://github.com/python-openxml/python-docx) as I needed features and quality of life improvements.

Key differences at a glance:
- Supporting multiple numbered lists within a document
- Supporting TOC updates within the package without the need to open the document manually
- Supporting floating images within documents
- Supporting the ability to transform word documents into PDF's
- Horizontal rules + paragraph bounding boxes
- External hyperlinks

## Installation

```
pip install skelmis-docx
```

## Example

```python
>>> from docx import Document

>>> document = Document()
>>> document.add_paragraph("It was a dark and stormy night.")
<docx.text.paragraph.Paragraph object at 0x10f19e760>
>>> document.save("dark-and-stormy.docx")

>>> document = Document("dark-and-stormy.docx")
>>> document.paragraphs[0].text
'It was a dark and stormy night.'
```

More information is available in the [documentation](https://skelmis-docx.readthedocs.io/en/latest/)
