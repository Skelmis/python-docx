"""Provides objects that can characterize image streams.

That characterization is as to content type and size, as a required step in including
them in a document.
"""

from skelmis.docx.image.bmp import Bmp
from skelmis.docx.image.gif import Gif
from skelmis.docx.image.jpeg import Exif, Jfif
from skelmis.docx.image.png import Png
from skelmis.docx.image.tiff import Tiff

SIGNATURES = (
    # class, offset, signature_bytes
    (Png, 0, b"\x89PNG\x0d\x0a\x1a\x0a"),
    (Jfif, 6, b"JFIF"),
    (Exif, 6, b"Exif"),
    (Gif, 0, b"GIF87a"),
    (Gif, 0, b"GIF89a"),
    (Tiff, 0, b"MM\x00*"),  # big-endian (Motorola) TIFF
    (Tiff, 0, b"II*\x00"),  # little-endian (Intel) TIFF
    (Bmp, 0, b"BM"),
)
