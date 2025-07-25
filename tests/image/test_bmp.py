"""Test suite for skelmis.docx.image.bmp module."""

import io

import pytest

from skelmis.docx.image.bmp import Bmp
from skelmis.docx.image.constants import MIME_TYPE

from ..unitutil.mock import ANY, initializer_mock


class DescribeBmp:
    def it_can_construct_from_a_bmp_stream(self, Bmp__init__):
        cx, cy, horz_dpi, vert_dpi = 26, 43, 200, 96
        bytes_ = (
            b"fillerfillerfiller\x1a\x00\x00\x00\x2b\x00\x00\x00"
            b"fillerfiller\xb8\x1e\x00\x00\x00\x00\x00\x00"
        )
        stream = io.BytesIO(bytes_)

        bmp = Bmp.from_stream(stream)

        Bmp__init__.assert_called_once_with(ANY, cx, cy, horz_dpi, vert_dpi)
        assert isinstance(bmp, Bmp)

    def it_knows_its_content_type(self):
        bmp = Bmp(None, None, None, None)
        assert bmp.content_type == MIME_TYPE.BMP

    def it_knows_its_default_ext(self):
        bmp = Bmp(None, None, None, None)
        assert bmp.default_ext == "bmp"

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def Bmp__init__(self, request):
        return initializer_mock(request, Bmp)
