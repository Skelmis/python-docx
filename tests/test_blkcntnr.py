"""Test suite for the skelmis.docx.blkcntnr (block item container) module."""

import pytest

from skelmis.docx import Document
from skelmis.docx.blkcntnr import BlockItemContainer
from skelmis.docx.shared import Inches
from skelmis.docx.table import Table
from skelmis.docx.text.paragraph import Paragraph

from .unitutil.cxml import element, xml
from .unitutil.file import snippet_seq, test_file
from .unitutil.mock import call, instance_mock, method_mock


class DescribeBlockItemContainer:
    """Unit-test suite for `docx.blkcntnr.BlockItemContainer`."""

    def it_can_add_a_paragraph(self, add_paragraph_fixture, _add_paragraph_):
        text, style, paragraph_, add_run_calls = add_paragraph_fixture
        _add_paragraph_.return_value = paragraph_
        blkcntnr = BlockItemContainer(None, None)

        paragraph = blkcntnr.add_paragraph(text, style)

        _add_paragraph_.assert_called_once_with(blkcntnr)
        assert paragraph.add_run.call_args_list == add_run_calls
        assert paragraph.style == style
        assert paragraph is paragraph_

    def it_can_add_a_table(self, add_table_fixture):
        blkcntnr, rows, cols, width, expected_xml = add_table_fixture
        table = blkcntnr.add_table(rows, cols, width)
        assert isinstance(table, Table)
        assert table._element.xml == expected_xml
        assert table._parent is blkcntnr

    def it_can_iterate_its_inner_content(self):
        document = Document(test_file("blk-inner-content.docx"))

        inner_content = document.iter_inner_content()

        para = next(inner_content)
        assert isinstance(para, Paragraph)
        assert para.text == "P1"
        # --
        t = next(inner_content)
        assert isinstance(t, Table)
        assert t.rows[0].cells[0].text == "T2"
        # --
        para = next(inner_content)
        assert isinstance(para, Paragraph)
        assert para.text == "P3"
        # --
        with pytest.raises(StopIteration):
            next(inner_content)

    def it_provides_access_to_the_paragraphs_it_contains(self, paragraphs_fixture):
        # test len(), iterable, and indexed access
        blkcntnr, expected_count = paragraphs_fixture
        paragraphs = blkcntnr.paragraphs
        assert len(paragraphs) == expected_count
        count = 0
        for idx, paragraph in enumerate(paragraphs):
            assert isinstance(paragraph, Paragraph)
            assert paragraphs[idx] is paragraph
            count += 1
        assert count == expected_count

    def it_provides_access_to_the_tables_it_contains(self, tables_fixture):
        # test len(), iterable, and indexed access
        blkcntnr, expected_count = tables_fixture
        tables = blkcntnr.tables
        assert len(tables) == expected_count
        count = 0
        for idx, table in enumerate(tables):
            assert isinstance(table, Table)
            assert tables[idx] is table
            count += 1
        assert count == expected_count

    def it_adds_a_paragraph_to_help(self, _add_paragraph_fixture):
        blkcntnr, expected_xml = _add_paragraph_fixture
        new_paragraph = blkcntnr._add_paragraph()
        assert isinstance(new_paragraph, Paragraph)
        assert new_paragraph._parent == blkcntnr
        assert blkcntnr._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("", None),
            ("Foo", None),
            ("", "Bar"),
            ("Foo", "Bar"),
        ]
    )
    def add_paragraph_fixture(self, request, paragraph_):
        text, style = request.param
        paragraph_.style = None
        add_run_calls = [call(text)] if text else []
        return text, style, paragraph_, add_run_calls

    @pytest.fixture
    def _add_paragraph_fixture(self, request):
        blkcntnr_cxml, after_cxml = "w:body", "w:body/w:p"
        blkcntnr = BlockItemContainer(element(blkcntnr_cxml), None)
        expected_xml = xml(after_cxml)
        return blkcntnr, expected_xml

    @pytest.fixture
    def add_table_fixture(self):
        blkcntnr = BlockItemContainer(element("w:body"), None)
        rows, cols, width = 2, 2, Inches(2)
        expected_xml = snippet_seq("new-tbl")[0]
        return blkcntnr, rows, cols, width, expected_xml

    @pytest.fixture(
        params=[
            ("w:body", 0),
            ("w:body/w:p", 1),
            ("w:body/(w:p,w:p)", 2),
            ("w:body/(w:p,w:tbl)", 1),
            ("w:body/(w:p,w:tbl,w:p)", 2),
        ]
    )
    def paragraphs_fixture(self, request):
        blkcntnr_cxml, expected_count = request.param
        blkcntnr = BlockItemContainer(element(blkcntnr_cxml), None)
        return blkcntnr, expected_count

    @pytest.fixture(
        params=[
            ("w:body", 0),
            ("w:body/w:tbl", 1),
            ("w:body/(w:tbl,w:tbl)", 2),
            ("w:body/(w:p,w:tbl)", 1),
            ("w:body/(w:tbl,w:tbl,w:p)", 2),
        ]
    )
    def tables_fixture(self, request):
        blkcntnr_cxml, expected_count = request.param
        blkcntnr = BlockItemContainer(element(blkcntnr_cxml), None)
        return blkcntnr, expected_count

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _add_paragraph_(self, request):
        return method_mock(request, BlockItemContainer, "_add_paragraph")

    @pytest.fixture
    def paragraph_(self, request):
        return instance_mock(request, Paragraph)
