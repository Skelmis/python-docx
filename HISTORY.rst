.. :changelog:

Release History
---------------

2.4.0 (2025-07-08)
++++++++++++++++++

* If a bookmark is created with text, return a `Run` object so the text can be edited / styled

2.3.0 (2025-06-15)
++++++++++++++++++

* Implement support to add bookmarks in a paragraph
* Implement support for creating internal hyperlinks to bookmarks

2.2.2 (2025-06-15)
++++++++++++++++++

* Resolve a Hatch build issue where the default `document.docx` file was not shipped in builds

2.2.1 (2025-06-15)
++++++++++++++++++

- Moves project from Poetry to uv
- Re-add's relevant pypi homepage links

2.2.0 (2025-06-15)
++++++++++++++++++

- Implement support for Table of Contents

2.1.0 (2025-03-01)
++++++++++++++++++

- Fix items not being fully moved to new namespace

2.0.0 (2025-02-16)
++++++++++++++++++

- Move to private `skelmis` namespace to reduce collision issues. `PR here <https://github.com/Skelmis/python-docx/pull/15>`

1.2.X -> 1.6.X
++++++++++++++

- Supporting multiple numbered lists within a document (`1 <https://skelmis-docx.readthedocs.io/en/latest/api/text.html#docx.text.paragraph.Paragraph.restart_numbering>`, `2 <https://skelmis-docx.readthedocs.io/en/latest/api/document.html#docx.document.Document.configure_styles_for_numbered_lists>`)
- Supporting TOC updates within the package without the need to open the document manually (`1 <https://skelmis-docx.readthedocs.io/en/latest/api/utility.html#docx.utility.update_toc>`, `2 <https://skelmis-docx.readthedocs.io/en/latest/api/utility.html#docx.utility.export_libre_macro>`)
- Supporting floating images within documents (`1 <https://skelmis-docx.readthedocs.io/en/latest/api/text.html#docx.text.run.Run.add_float_picture>`)
- Supporting the ability to transform word documents into PDF's (`1 <https://skelmis-docx.readthedocs.io/en/latest/api/utility.html#docx.utility.document_to_pdf>`)
- Horizontal rules + paragraph bounding boxes / borders (`1 <https://skelmis-docx.readthedocs.io/en/latest/api/text.html#docx.text.paragraph.Paragraph.insert_horizontal_rule>`, `2 <https://skelmis-docx.readthedocs.io/en/latest/api/text.html#docx.text.paragraph.Paragraph.draw_paragraph_border>`)
- External hyperlinks (`1 <https://skelmis-docx.readthedocs.io/en/latest/api/text.html#docx.text.paragraph.Paragraph.add_external_hyperlink>`)

1.1.2 (2024-05-01)
++++++++++++++++++

- Fix #1383 Revert lxml<=4.9.2 pin that breaks Python 3.12 install
- Fix #1385 Support use of Part._rels by python-docx-template
- Add support and testing for Python 3.12

1.1.1 (2024-04-29)
++++++++++++++++++

- Fix #531, #1146 Index error on table with misaligned borders
- Fix #1335 Tolerate invalid float value in bottom-margin
- Fix #1337 Do not require typing-extensions at runtime


1.1.0 (2023-11-03)
++++++++++++++++++

- Add BlockItemContainer.iter_inner_content()


1.0.1 (2023-10-12)
++++++++++++++++++

- Fix #1256: parse_xml() and OxmlElement moved.
- Add Hyperlink.fragment and .url


1.0.0 (2023-10-01)
+++++++++++++++++++

- Remove Python 2 support. Supported versions are 3.7+
- Fix #85:   Paragraph.text includes hyperlink text
- Add #1113: Hyperlink.address
- Add Hyperlink.contains_page_break
- Add Hyperlink.runs
- Add Hyperlink.text
- Add Paragraph.contains_page_break
- Add Paragraph.hyperlinks
- Add Paragraph.iter_inner_content()
- Add Paragraph.rendered_page_breaks
- Add RenderedPageBreak.following_paragraph_fragment
- Add RenderedPageBreak.preceding_paragraph_fragment
- Add Run.contains_page_break
- Add Run.iter_inner_content()
- Add Section.iter_inner_content()


0.8.11 (2021-05-15)
+++++++++++++++++++

- Small build changes and Python 3.8 version changes like collections.abc location.


0.8.10 (2019-01-08)
+++++++++++++++++++

- Revert use of expanded package directory for default.docx to work around setup.py
  problem with filenames containing square brackets.


0.8.9 (2019-01-08)
++++++++++++++++++

- Fix gap in MANIFEST.in that excluded default document template directory


0.8.8 (2019-01-07)
++++++++++++++++++

- Add support for headers and footers


0.8.7 (2018-08-18)
++++++++++++++++++

- Add _Row.height_rule
- Add _Row.height
- Add _Cell.vertical_alignment
- Fix #455: increment next_id, don't fill gaps
- Add #375: import docx failure on --OO optimization
- Add #254: remove default zoom percentage
- Add #266: miscellaneous documentation fixes
- Add #175: refine MANIFEST.ini
- Add #168: Unicode error on core-props in Python 2


0.8.6 (2016-06-22)
++++++++++++++++++

- Add #257: add Font.highlight_color
- Add #261: add ParagraphFormat.tab_stops
- Add #303: disallow XML entity expansion


0.8.5 (2015-02-21)
++++++++++++++++++

- Fix #149: KeyError on Document.add_table()
- Fix #78: feature: add_table() sets cell widths
- Add #106: feature: Table.direction (i.e. right-to-left)
- Add #102: feature: add CT_Row.trPr


0.8.4 (2015-02-20)
++++++++++++++++++

- Fix #151: tests won't run on PyPI distribution
- Fix #124: default to inches on no TIFF resolution unit


0.8.3 (2015-02-19)
++++++++++++++++++

- Add #121, #135, #139: feature: Font.color


0.8.2 (2015-02-16)
++++++++++++++++++

- Fix #94: picture prints at wrong size when scaled
- Extract `docx.document.Document` object from `DocumentPart`

  Refactor `docx.Document` from an object into a factory function for new
  `docx.document.Document object`. Extract methods from prior `docx.Document`
  and `docx.parts.document.DocumentPart` to form the new API class and retire
  `docx.Document` class.

- Migrate `Document.numbering_part` to `DocumentPart.numbering_part`. The
  `numbering_part` property is not part of the published API and is an
  interim internal feature to be replaced in a future release, perhaps with
  something like `Document.numbering_definitions`. In the meantime, it can
  now be accessed using ``Document.part.numbering_part``.


0.8.1 (2015-02-10)
++++++++++++++++++

- Fix #140: Warning triggered on Document.add_heading/table()


0.8.0 (2015-02-08)
++++++++++++++++++

- Add styles. Provides general capability to access and manipulate paragraph,
  character, and table styles.

- Add ParagraphFormat object, accessible on Paragraph.paragraph_format, and
  providing the following paragraph formatting properties:

  + paragraph alignment (justfification)
  + space before and after paragraph
  + line spacing
  + indentation
  + keep together, keep with next, page break before, and widow control

- Add Font object, accessible on Run.font, providing character-level
  formatting including:

  + typeface (e.g. 'Arial')
  + point size
  + underline
  + italic
  + bold
  + superscript and subscript

The following issues were retired:

- Add feature #56: superscript/subscript
- Add feature #67: lookup style by UI name
- Add feature #98: Paragraph indentation
- Add feature #120: Document.styles

**Backward incompatibilities**

Paragraph.style now returns a Style object. Previously it returned the style
name as a string. The name can now be retrieved using the Style.name
property, for example, `paragraph.style.name`.


0.7.6 (2014-12-14)
++++++++++++++++++

- Add feature #69: Table.alignment
- Add feature #29: Document.core_properties


0.7.5 (2014-11-29)
++++++++++++++++++

- Add feature #65: _Cell.merge()


0.7.4 (2014-07-18)
++++++++++++++++++

- Add feature #45: _Cell.add_table()
- Add feature #76: _Cell.add_paragraph()
- Add _Cell.tables property (read-only)


0.7.3 (2014-07-14)
++++++++++++++++++

- Add Table.autofit
- Add feature #46: _Cell.width


0.7.2 (2014-07-13)
++++++++++++++++++

- Fix: Word does not interpret <w:cr/> as line feed


0.7.1 (2014-07-11)
++++++++++++++++++

- Add feature #14: Run.add_picture()


0.7.0 (2014-06-27)
++++++++++++++++++

- Add feature #68: Paragraph.insert_paragraph_before()
- Add feature #51: Paragraph.alignment (read/write)
- Add feature #61: Paragraph.text setter
- Add feature #58: Run.add_tab()
- Add feature #70: Run.clear()
- Add feature #60: Run.text setter
- Add feature #39: Run.text and Paragraph.text interpret '\n' and '\t' chars


0.6.0 (2014-06-22)
++++++++++++++++++

- Add feature #15: section page size
- Add feature #66: add section
- Add page margins and page orientation properties on Section
- Major refactoring of oxml layer


0.5.3 (2014-05-10)
++++++++++++++++++

- Add feature #19: Run.underline property


0.5.2 (2014-05-06)
++++++++++++++++++

- Add feature #17: character style


0.5.1 (2014-04-02)
++++++++++++++++++

- Fix issue #23, `Document.add_picture()` raises ValueError when document
  contains VML drawing.


0.5.0 (2014-03-02)
++++++++++++++++++

- Add 20 tri-state properties on Run, including all-caps, double-strike,
  hidden, shadow, small-caps, and 15 others.


0.4.0 (2014-03-01)
++++++++++++++++++

- Advance from alpha to beta status.
- Add pure-python image header parsing; drop Pillow dependency


0.3.0a5 (2014-01-10)
++++++++++++++++++++++

- Hotfix: issue #4, Document.add_picture() fails on second and subsequent
  images.


0.3.0a4 (2014-01-07)
++++++++++++++++++++++

- Complete Python 3 support, tested on Python 3.3


0.3.0a3 (2014-01-06)
++++++++++++++++++++++

- Fix setup.py error on some Windows installs


0.3.0a1 (2014-01-05)
++++++++++++++++++++++

- Full object-oriented rewrite
- Feature-parity with prior version
- text: add paragraph, run, text, bold, italic
- table: add table, add row, add column
- styles: specify style for paragraph, table
- picture: add inline picture, auto-scaling
- breaks: add page break
- tests: full pytest and behave-based 2-layer test suite


0.3.0dev1 (2013-12-14)
++++++++++++++++++++++

- Round-trip .docx file, preserving all parts and relationships
- Load default "template" .docx on open with no filename
- Open from stream and save to stream (file-like object)
- Add paragraph at and of document
