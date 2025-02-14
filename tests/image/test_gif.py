"""Unit test suite for docx.image.gif module."""

import io

import pytest

from skelmis.docx.image.constants import MIME_TYPE
from skelmis.docx.image.gif import Gif

from ..unitutil.mock import ANY, initializer_mock


class DescribeGif:
    def it_can_construct_from_a_gif_stream(self, Gif__init__):
        cx, cy = 42, 24
        bytes_ = b"filler\x2A\x00\x18\x00"
        stream = io.BytesIO(bytes_)

        gif = Gif.from_stream(stream)

        Gif__init__.assert_called_once_with(ANY, cx, cy, 72, 72)
        assert isinstance(gif, Gif)

    def it_knows_its_content_type(self):
        gif = Gif(None, None, None, None)
        assert gif.content_type == MIME_TYPE.GIF

    def it_knows_its_default_ext(self):
        gif = Gif(None, None, None, None)
        assert gif.default_ext == "gif"

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Gif__init__(self, request):
        return initializer_mock(request, Gif)
