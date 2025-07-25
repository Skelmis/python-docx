# pyright: reportImportCycles=false
# pyright: reportPrivateUsage=false

"""|Document| and closely related objects."""

from __future__ import annotations

from pathlib import Path
from typing import IO, TYPE_CHECKING, Iterator, List

import skelmis.docx
from skelmis.docx.blkcntnr import BlockItemContainer
from skelmis.docx.enum.section import WD_SECTION
from skelmis.docx.enum.text import WD_BREAK
from skelmis.docx.oxml import simpletypes
from skelmis.docx.section import Section, Sections
from skelmis.docx.shared import ElementProxy, Emu

if TYPE_CHECKING:
    import skelmis.docx.types as t
    from skelmis.docx.oxml.document import CT_Body, CT_Document
    from skelmis.docx.parts.document import DocumentPart
    from skelmis.docx.settings import Settings
    from skelmis.docx.shared import Length
    from skelmis.docx.styles.style import ParagraphStyle, _TableStyle
    from skelmis.docx.table import Table
    from skelmis.docx.text.paragraph import Paragraph


class Document(ElementProxy):
    """WordprocessingML (WML) document.

    Not intended to be constructed directly. Use :func:`docx.Document` to open or create
    a document.
    """

    def __init__(self, element: CT_Document, part: DocumentPart):
        super(Document, self).__init__(element)
        self._element = element
        self._part = part
        self.__body = None

    def configure_styles_for_numbered_lists(self):
        """Configures the underlying document such that you
        can include multiple numbered lists with correct numbers.

        If you wish to change the appearance of the resultant styles
        then you should override this method with your own styling choices
        as these are shipped 'as is' and are generally good enough.
        """
        STYP = skelmis.docx.enum.style.WD_STYLE_TYPE
        num_xml = self.part.numbering_part.element
        next_abstract_id = max([J.abstractNumId for J in num_xml.abstractNum_lst]) + 1
        l = num_xml._new_abstractNum()  # noqa: E741
        l.abstractNumId = next_abstract_id
        l.add_multiLevelType().val = "multilevel"

        formats = {
            0: "decimal",
            1: "decimal",
            2: "decimal",
        }
        text_fmts = {
            0: "%1.",
            1: "%1.%2.",
            2: "%1.%2.%3.",
        }
        starts = {0: 1, 1: 1, 2: 1}
        restarts = {0: False, 1: False, 2: 1}
        hosts = {0: "List Number", 1: "List Number 2", 2: "List Number 3"}

        num_xml.abstractNum_lst[-1].addnext(l)
        nNum = num_xml.add_num(next_abstract_id)

        for i in range(3):
            lvl = l.add_lvl()
            lvl.ilvl = i
            lvl.add_start().val = starts[i]
            lvl.add_numFmt().val = formats[i]
            if restarts[i]:
                lvl.add_lvlRestart().val = restarts[i]
            lvl.add_lvlText().val = text_fmts[i]
            lvl.add_suff().val = "tab"
            p_pr = lvl.add_pPr()
            p_pr.ind_left = simpletypes.Twips(i * 720)
            ho = self.styles.get_by_id(
                self.styles.get_style_id(hosts[i], STYP.PARAGRAPH), STYP.PARAGRAPH
            ).element.pPr.numPr
            ho.get_or_add_ilvl().val = i
            ho.get_or_add_numId().val = nNum.numId

    def add_heading(self, text: str = "", level: int = 1):
        """Return a heading paragraph newly added to the end of the document.

        The heading paragraph will contain `text` and have its paragraph style
        determined by `level`. If `level` is 0, the style is set to `Title`. If `level`
        is 1 (or omitted), `Heading 1` is used. Otherwise the style is set to `Heading
        {level}`. Raises |ValueError| if `level` is outside the range 0-9.
        """
        if not 0 <= level <= 9:
            raise ValueError("level must be in range 0-9, got %d" % level)
        style = "Title" if level == 0 else "Heading %d" % level
        return self.add_paragraph(text, style)

    def add_page_break(self):
        """Return newly |Paragraph| object containing only a page break."""
        paragraph = self.add_paragraph()
        paragraph.add_run().add_break(WD_BREAK.PAGE)
        return paragraph

    def add_paragraph(self, text: str = "", style: str | ParagraphStyle | None = None) -> Paragraph:
        """Return paragraph newly added to the end of the document.

        The paragraph is populated with `text` and having paragraph style `style`.

        `text` can contain tab (``\\t``) characters, which are converted to the
        appropriate XML form for a tab. `text` can also include newline (``\\n``) or
        carriage return (``\\r``) characters, each of which is converted to a line
        break.
        """
        return self._body.add_paragraph(text, style)

    def add_floating_picture(self):
        """"""
        # Source reference: https://github.com/ArtifexSoftware/pdf2docx/issues/54#issuecomment-715925252

    def add_picture(
        self,
        image_path_or_stream: str | IO[bytes],
        width: int | Length | None = None,
        height: int | Length | None = None,
    ):
        """Return new picture shape added in its own paragraph at end of the document.

        The picture contains the image at `image_path_or_stream`, scaled based on
        `width` and `height`. If neither width nor height is specified, the picture
        appears at its native size. If only one is specified, it is used to compute a
        scaling factor that is then applied to the unspecified dimension, preserving the
        aspect ratio of the image. The native size of the picture is calculated using
        the dots-per-inch (dpi) value specified in the image file, defaulting to 72 dpi
        if no value is specified, as is often the case.
        """
        run = self.add_paragraph().add_run()
        return run.add_picture(image_path_or_stream, width, height)

    def add_section(self, start_type: WD_SECTION = WD_SECTION.NEW_PAGE):
        """Return a |Section| object newly added at the end of the document.

        The optional `start_type` argument must be a member of the :ref:`WdSectionStart`
        enumeration, and defaults to ``WD_SECTION.NEW_PAGE`` if not provided.
        """
        new_sectPr = self._element.body.add_section_break()
        new_sectPr.start_type = start_type
        return Section(new_sectPr, self._part)

    def add_table(self, rows: int, cols: int, style: str | _TableStyle | None = None):
        """Add a table having row and column counts of `rows` and `cols` respectively.

        `style` may be a table style object or a table style name. If `style` is |None|,
        the table inherits the default table style of the document.
        """
        table = self._body.add_table(rows, cols, self._block_width)
        table.style = style
        return table

    @property
    def core_properties(self):
        """A |CoreProperties| object providing Dublin Core properties of document."""
        return self._part.core_properties

    @property
    def inline_shapes(self):
        """The |InlineShapes| collection for this document.

        An inline shape is a graphical object, such as a picture, contained in a run of
        text and behaving like a character glyph, being flowed like other text in a
        paragraph.
        """
        return self._part.inline_shapes

    def iter_inner_content(self) -> Iterator[Paragraph | Table]:
        """Generate each `Paragraph` or `Table` in this document in document order."""
        return self._body.iter_inner_content()

    @property
    def paragraphs(self) -> List[Paragraph]:
        """The |Paragraph| instances in the document, in document order.

        Note that paragraphs within revision marks such as ``<w:ins>`` or ``<w:del>`` do
        not appear in this list.
        """
        return self._body.paragraphs

    @property
    def part(self) -> DocumentPart:
        """The |DocumentPart| object of this document."""
        return self._part

    def save(self, path_or_stream: str | Path | IO[bytes]):
        """Save this document to `path_or_stream`.

        `path_or_stream` can be either a path to a filesystem location (a string) or a
        file-like object.
        """
        if isinstance(path_or_stream, Path):
            path_or_stream = str(path_or_stream)

        self._part.save(path_or_stream)

    @property
    def sections(self) -> Sections:
        """|Sections| object providing access to each section in this document."""
        return Sections(self._element, self._part)

    @property
    def settings(self) -> Settings:
        """A |Settings| object providing access to the document-level settings."""
        return self._part.settings

    @property
    def styles(self):
        """A |Styles| object providing access to the styles in this document."""
        return self._part.styles

    @property
    def tables(self) -> List[Table]:
        """All |Table| instances in the document, in document order.

        Note that only tables appearing at the top level of the document appear in this
        list; a table nested inside a table cell does not appear. A table within
        revision marks such as ``<w:ins>`` or ``<w:del>`` will also not appear in the
        list.
        """
        return self._body.tables

    @property
    def _block_width(self) -> Length:
        """A |Length| object specifying the space between margins in last section."""
        section = self.sections[-1]
        return Emu(section.page_width - section.left_margin - section.right_margin)

    @property
    def _body(self) -> _Body:
        """The |_Body| instance containing the content for this document."""
        if self.__body is None:
            self.__body = _Body(self._element.body, self)
        return self.__body


class _Body(BlockItemContainer):
    """Proxy for `<w:body>` element in this document.

    It's primary role is a container for document content.
    """

    def __init__(self, body_elm: CT_Body, parent: t.ProvidesStoryPart):
        super(_Body, self).__init__(body_elm, parent)
        self._body = body_elm

    def clear_content(self):
        """Return this |_Body| instance after clearing it of all content.

        Section properties for the main document story, if present, are preserved.
        """
        self._body.clear_content()
        return self
