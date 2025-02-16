.. _WdStyleType:

``WD_STYLE_TYPE``
=================

Specifies one of the four style types: paragraph, character, list, or
table.

Example::

    from skelmis.docx import Document
    from skelmis.docx.enum.style import WD_STYLE_TYPE

    styles = Document().styles
    assert styles[0].type == WD_STYLE_TYPE.PARAGRAPH

----

CHARACTER
    Character style.

LIST
    List style.

PARAGRAPH
    Paragraph style.

TABLE
    Table style.
