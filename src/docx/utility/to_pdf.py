import secrets
import shutil
import subprocess
import sys
from pathlib import Path


def _update_toc_linux(docx_file: Path) -> None:
    """TOC bindings for linux"""
    # This method hangs if item is already open so we cheat a little here
    tmp_file = str(docx_file) + f".{secrets.token_hex(4)}.docx"
    tmp_file = Path(tmp_file)
    shutil.copy(docx_file, tmp_file)

    # Source: https://github.com/python-openxml/python-docx/issues/1207#issuecomment-1924053420
    subprocess.call(
        [
            "libreoffice",
            "--headless",
            f"macro:///Standard.Module1.UpdateTOC({str(tmp_file)})",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    shutil.copy(tmp_file, docx_file)
    tmp_file.unlink()


def export_libre_macro(
    macro_folder: Path = Path("~/.config/libreoffice/4/user/basic/Standard"),
) -> None:
    """Automatically moves the LibreOffice macro file to `macro_folder`.

    Warning, this overrides Module1.xba

    :py:class:`Path` is where your macros live
    """
    macro_folder = macro_folder.expanduser()
    module_file = Path(__file__).parent.resolve() / "Module1.xba"
    shutil.copy(module_file, macro_folder)


def update_toc(docx_file: Path | str) -> None:
    """Update a TOC within a word document.

    If you are on linux, please call `export_libre_macro` first.
    """
    if isinstance(docx_file, str):
        docx_file = Path(docx_file)

    docx_file = docx_file.absolute()

    if sys.platform == "linux":
        _update_toc_linux(docx_file)
    elif sys.platform == "win32":
        raise ValueError("Windows is not yet implemented yet. Consider docx2pdf for now")
    else:
        raise ValueError(f"{sys.platform} is not implemented")
