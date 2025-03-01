"""Paragraph-related proxy types."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List, cast, Literal

from skelmis.docx.enum.style import WD_STYLE_TYPE
from skelmis.docx.oxml import OxmlElement
from skelmis.docx.oxml.ns import qn
from skelmis.docx.oxml.text.run import CT_R
from skelmis.docx.shared import StoryChild
from skelmis.docx.styles.style import ParagraphStyle
from skelmis.docx.text.hyperlink import Hyperlink
from skelmis.docx.text.pagebreak import RenderedPageBreak
from skelmis.docx.text.parfmt import ParagraphFormat
from skelmis.docx.text.run import Run
from skelmis.docx.opc.constants import RELATIONSHIP_TYPE

if TYPE_CHECKING:
    import skelmis.docx.types as t
    from skelmis.docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from skelmis.docx.oxml.text.paragraph import CT_P
    from skelmis.docx.styles.style import CharacterStyle


class Paragraph(StoryChild):
    """Proxy object wrapping a `<w:p>` element."""

    def __init__(self, p: CT_P, parent: t.ProvidesStoryPart):
        super(Paragraph, self).__init__(parent)
        self._p = self._element = p

    def add_external_hyperlink(
        self,
        url: str,
        text: str,
        *,
        color: str | None = "0000FF",
        underline: bool = True,
    ) -> Hyperlink:
        """
        A function that places an external hyperlink within a paragraph object.

        Default behaviour is Blue with underlined text.

        :param url: A string containing the required url
        :param text: The text displayed for the url
        :param color: The color of the text displayed
        :param underline: Whether the text is underlined or not
        :return: The hyperlink object.
        """
        # Sourced from https://github.com/python-openxml/python-docx/issues/74#issuecomment-261169410

        # This gets access to the document.xml.rels file and gets a new relation id value
        part = self.part
        r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

        # Create the w:hyperlink tag and add needed values
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(
            qn("r:id"),
            r_id,
        )

        # Create a w:r element
        new_run = OxmlElement("w:r")

        # Create a new w:rPr element
        rPr = OxmlElement("w:rPr")

        # Add color if it is given
        if not color is None:
            c = OxmlElement("w:color")
            c.set(qn("w:val"), color)
            rPr.append(c)

        if underline:
            u = OxmlElement("w:u")
            u.set(qn("w:val"), "single")
            rPr.append(u)
        else:
            u = OxmlElement("w:u")
            u.set(qn("w:val"), "none")
            rPr.append(u)

        # Join all the xml elements together and add the required text to the w:r element
        new_run.append(rPr)
        new_run.text = text
        hyperlink.append(new_run)

        self._p.append(hyperlink)
        return Hyperlink(hyperlink, self)

    def insert_horizontal_rule(self):
        """Insert a horizontal rule at the bottom of the current paragraph."""
        self._draw_bounding_line(bottom=True)

    def draw_paragraph_border(
        self,
        *,
        top: bool = False,
        bottom: bool = False,
        right: bool = False,
        left: bool = False,
    ):
        """Draw's a line around the current paragraph corresponding the provided arguments.

        Valid arguments are top, bottom, right left. All off by default.
        """
        self._draw_bounding_line(top=top, bottom=bottom, right=right, left=left)

    # noinspection DuplicatedCode
    def _draw_bounding_line(
        self,
        *,
        top: bool = False,
        bottom: bool = False,
        right: bool = False,
        left: bool = False,
    ):
        # Original sources:
        # - https://stackoverflow.com/a/68530806/13781503
        # - https://github.com/python-openxml/python-docx/issues/105
        p = self._p  # p is the <w:p> XML element
        pPr = p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        pPr.insert_element_before(
            pBdr,
            "w:shd",
            "w:tabs",
            "w:suppressAutoHyphens",
            "w:kinsoku",
            "w:wordWrap",
            "w:overflowPunct",
            "w:topLinePunct",
            "w:autoSpaceDE",
            "w:autoSpaceDN",
            "w:bidi",
            "w:adjustRightInd",
            "w:snapToGrid",
            "w:spacing",
            "w:ind",
            "w:contextualSpacing",
            "w:mirrorIndents",
            "w:suppressOverlap",
            "w:jc",
            "w:textDirection",
            "w:textAlignment",
            "w:textboxTightWrap",
            "w:outlineLvl",
            "w:divId",
            "w:cnfStyle",
            "w:rPr",
            "w:sectPr",
            "w:pPrChange",
        )
        if top:
            top_s = OxmlElement(f"w:top")
            top_s.set(qn("w:val"), "single")
            top_s.set(qn("w:sz"), "6")
            top_s.set(qn("w:space"), "1")
            top_s.set(qn("w:color"), "auto")
            pBdr.append(top_s)

        if bottom:
            bottom_s = OxmlElement(f"w:bottom")
            bottom_s.set(qn("w:val"), "single")
            bottom_s.set(qn("w:sz"), "6")
            bottom_s.set(qn("w:space"), "1")
            bottom_s.set(qn("w:color"), "auto")
            pBdr.append(bottom_s)

        if left:
            left_s = OxmlElement(f"w:left")
            left_s.set(qn("w:val"), "single")
            left_s.set(qn("w:sz"), "6")
            left_s.set(qn("w:space"), "1")
            left_s.set(qn("w:color"), "auto")
            pBdr.append(left_s)

        if right:
            right_s = OxmlElement(f"w:right")
            right_s.set(qn("w:val"), "single")
            right_s.set(qn("w:sz"), "6")
            right_s.set(qn("w:space"), "1")
            right_s.set(qn("w:color"), "auto")
            pBdr.append(right_s)

    def restart_numbering(self):
        """
        Restarting the numbering of paragraph

        Raises ValueError if you call this on a
        paragraph which does not contain a numbered list.
        """

        # Getting the abstract number of paragraph
        try:
            abstract_num_id = self.part.document.part.numbering_part.element.num_having_numId(
                self.style.element.get_or_add_pPr().get_or_add_numPr().numId.val
            ).abstractNumId.val
        except AttributeError as e:
            raise ValueError(
                "Are you sure this paragraph contains a numbered list? It doesn't appear so."
            ) from e

        # Add abstract number to numbering part and reset
        num = self.part.numbering_part.element.add_num(abstract_num_id)
        num.add_lvlOverride(ilvl=0).add_startOverride(1)

        # Get or add elements to paragraph
        p_pr = self._p.get_or_add_pPr()
        num_pr = p_pr.get_or_add_numPr()
        ilvl = num_pr.get_or_add_ilvl()
        ilvl.val = 0
        num_id = num_pr.get_or_add_numId()
        num_id.val = int(num.numId)

    def add_run(self, text: str | None = None, style: str | CharacterStyle | None = None) -> Run:
        """Append run containing `text` and having character-style `style`.

        `text` can contain tab (``\\t``) characters, which are converted to the
        appropriate XML form for a tab. `text` can also include newline (``\\n``) or
        carriage return (``\\r``) characters, each of which is converted to a line
        break. When `text` is `None`, the new run is empty.
        """
        r = self._p.add_r()
        run = Run(r, self)
        if text:
            run.text = text
        if style:
            run.style = style
        return run

    @property
    def alignment(self) -> WD_PARAGRAPH_ALIGNMENT | None:
        """A member of the :ref:`WdParagraphAlignment` enumeration specifying the
        justification setting for this paragraph.

        A value of |None| indicates the paragraph has no directly-applied alignment
        value and will inherit its alignment value from its style hierarchy. Assigning
        |None| to this property removes any directly-applied alignment value.
        """
        return self._p.alignment

    @alignment.setter
    def alignment(self, value: WD_PARAGRAPH_ALIGNMENT):
        self._p.alignment = value

    def clear(self):
        """Return this same paragraph after removing all its content.

        Paragraph-level formatting, such as style, is preserved.
        """
        self._p.clear_content()
        return self

    @property
    def contains_page_break(self) -> bool:
        """`True` when one or more rendered page-breaks occur in this paragraph."""
        return bool(self._p.lastRenderedPageBreaks)

    @property
    def hyperlinks(self) -> List[Hyperlink]:
        """A |Hyperlink| instance for each hyperlink in this paragraph."""
        return [Hyperlink(hyperlink, self) for hyperlink in self._p.hyperlink_lst]

    def insert_paragraph_before(
        self, text: str | None = None, style: str | ParagraphStyle | None = None
    ) -> Paragraph:
        """Return a newly created paragraph, inserted directly before this paragraph.

        If `text` is supplied, the new paragraph contains that text in a single run. If
        `style` is provided, that style is assigned to the new paragraph.
        """
        paragraph = self._insert_paragraph_before()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    def iter_inner_content(self) -> Iterator[Run | Hyperlink]:
        """Generate the runs and hyperlinks in this paragraph, in the order they appear.

        The content in a paragraph consists of both runs and hyperlinks. This method
        allows accessing each of those separately, in document order, for when the
        precise position of the hyperlink within the paragraph text is important. Note
        that a hyperlink itself contains runs.
        """
        for r_or_hlink in self._p.inner_content_elements:
            yield (
                Run(r_or_hlink, self)
                if isinstance(r_or_hlink, CT_R)
                else Hyperlink(r_or_hlink, self)
            )

    @property
    def paragraph_format(self):
        """The |ParagraphFormat| object providing access to the formatting properties
        for this paragraph, such as line spacing and indentation."""
        return ParagraphFormat(self._element)

    @property
    def rendered_page_breaks(self) -> List[RenderedPageBreak]:
        """All rendered page-breaks in this paragraph.

        Most often an empty list, sometimes contains one page-break, but can contain
        more than one is rare or contrived cases.
        """
        return [RenderedPageBreak(lrpb, self) for lrpb in self._p.lastRenderedPageBreaks]

    @property
    def runs(self) -> List[Run]:
        """Sequence of |Run| instances corresponding to the <w:r> elements in this
        paragraph."""
        return [Run(r, self) for r in self._p.r_lst]

    @property
    def style(self) -> ParagraphStyle | None:
        """Read/Write.

        |_ParagraphStyle| object representing the style assigned to this paragraph. If
        no explicit style is assigned to this paragraph, its value is the default
        paragraph style for the document. A paragraph style name can be assigned in lieu
        of a paragraph style object. Assigning |None| removes any applied style, making
        its effective value the default paragraph style for the document.
        """
        style_id = self._p.style
        style = self.part.get_style(style_id, WD_STYLE_TYPE.PARAGRAPH)
        return cast(ParagraphStyle, style)

    @style.setter
    def style(self, style_or_name: str | ParagraphStyle | None):
        style_id = self.part.get_style_id(style_or_name, WD_STYLE_TYPE.PARAGRAPH)
        self._p.style = style_id

    @property
    def text(self) -> str:
        """The textual content of this paragraph.

        The text includes the visible-text portion of any hyperlinks in the paragraph.
        Tabs and line breaks in the XML are mapped to ``\\t`` and ``\\n`` characters
        respectively.

        Assigning text to this property causes all existing paragraph content to be
        replaced with a single run containing the assigned text. A ``\\t`` character in
        the text is mapped to a ``<w:tab/>`` element and each ``\\n`` or ``\\r``
        character is mapped to a line break. Paragraph-level formatting, such as style,
        is preserved. All run-level formatting, such as bold or italic, is removed.
        """
        return self._p.text

    @text.setter
    def text(self, text: str | None):
        self.clear()
        self.add_run(text)

    def _insert_paragraph_before(self):
        """Return a newly created paragraph, inserted directly before this paragraph."""
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)
