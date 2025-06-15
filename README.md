# python-docx

*python-docx* is a Python library for reading, creating, and updating Microsoft Word 2007+ (.docx) files.

This repository exists as a fork of [the official repo](https://github.com/python-openxml/python-docx) as I needed features and quality of life improvements.

Key differences at a glance:
- Supporting multiple numbered lists within a document ([1](https://skelmis-docx.readthedocs.io/en/latest/api/text.html#skelmis.docx.text.paragraph.Paragraph.restart_numbering), [2](https://skelmis-docx.readthedocs.io/en/latest/api/document.html#skelmis.docx.document.Document.configure_styles_for_numbered_lists))
- Supporting TOC updates within the package without the need to open the document manually ([1](https://skelmis-docx.readthedocs.io/en/latest/api/utility.html#skelmis.docx.utility.update_toc), [2](https://skelmis-docx.readthedocs.io/en/latest/api/utility.html#skelmis.docx.utility.export_libre_macro))
- Supporting floating images within documents ([1](https://skelmis-docx.readthedocs.io/en/latest/api/text.html#skelmis.docx.text.run.Run.add_float_picture))
- Supporting the ability to transform word documents into PDF's ([1](https://skelmis-docx.readthedocs.io/en/latest/api/utility.html#skelmis.docx.utility.document_to_pdf))
- Horizontal rules + paragraph bounding boxes / borders ([1](https://skelmis-docx.readthedocs.io/en/latest/api/text.html#skelmis.docx.text.paragraph.Paragraph.insert_horizontal_rule), [2](https://skelmis-docx.readthedocs.io/en/latest/api/text.html#skelmis.docx.text.paragraph.Paragraph.draw_paragraph_border))
- External hyperlinks ([1](https://skelmis-docx.readthedocs.io/en/latest/api/text.html#skelmis.docx.text.paragraph.Paragraph.add_external_hyperlink))
- Internal hyperlinks (Linking to bookmarks) ([1](https://skelmis-docx.readthedocs.io/en/latest/api/text.html#skelmis.docx.text.paragraph.Paragraph.add_internal_hyperlink))
- Creating bookmarks ([1](https://skelmis-docx.readthedocs.io/en/latest/api/text.html#skelmis.docx.text.paragraph.Paragraph.add_bookmark))
- The ability to insert a customisable Table of Contents (ToC) ([1](https://skelmis-docx.readthedocs.io/en/latest/api/text.html#skelmis.docx.text.paragraph.Paragraph.insert_table_of_contents))

## Installation

```
pip install skelmis-docx
```

## Example

```python
>>> from skelmis.docx import Document

>>> document = Document()
>>> document.add_paragraph("It was a dark and stormy night.")
<docx.text.paragraph.Paragraph object at 0x10f19e760>
>>> document.save("dark-and-stormy.docx")

>>> document = Document("dark-and-stormy.docx")
>>> document.paragraphs[0].text
'It was a dark and stormy night.'
```

More information is available in the [documentation](https://skelmis-docx.readthedocs.io/en/latest/)
