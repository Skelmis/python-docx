import json
import logging
import secrets
import shutil
import subprocess
import sys
from pathlib import Path

log = logging.getLogger(__name__)


def _update_toc_linux(docx_file: Path) -> None:
    """TOC bindings for linux"""
    # This method hangs if item is already open, so we cheat a little here
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


def _create_pdf_windows(docx_file: Path) -> None:
    import win32com.client

    word = win32com.client.Dispatch("Word.Application")
    wdFormatPDF = 17

    docx_filepath = docx_file
    pdf_filepath = Path(f"{docx_file.stem}.pdf").resolve()
    doc = word.Documents.Open(str(docx_filepath))
    try:
        doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
    except:
        raise
    finally:
        doc.Close(0)

    word.Quit()


def _create_pdf_linux(docx_file: Path) -> None:
    subprocess.call(
        [
            "libreoffice",
            "--convert-to",
            "pdf",
            str(docx_file),
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def _create_pdf_macos(docx_file: Path) -> None:
    log.warning("DOCX -> PDF on mac is untested. Any issues please raise an issue.")
    script = (Path(__file__).parent / "convert.jxa").resolve()
    cmd = [
        "/usr/bin/osascript",
        "-l",
        "JavaScript",
        str(script),
        str(docx_file),
        str(Path(f"{docx_file.stem}.pdf").resolve()),
    ]

    process = subprocess.Popen(cmd, stderr=subprocess.PIPE)
    process.wait()
    if process.returncode != 0:
        msg = process.stderr.read().decode().rstrip()
        if "application can't be found" in msg.lower():
            raise EnvironmentError("Microsoft Word is not available.")
        raise RuntimeError(msg)

    def stderr_results(process):
        while True:
            output_line = process.stderr.readline().rstrip()
            if not output_line:
                break
            yield output_line.decode("utf-8")

    for line in stderr_results(process):
        try:
            msg = json.loads(line)
        except ValueError:
            continue
        if msg["result"] == "error":
            print(msg)
            sys.exit(1)


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

    docx_file = docx_file.resolve().absolute()

    if sys.platform == "linux":
        _update_toc_linux(docx_file)
    elif sys.platform == "win32":
        raise ValueError("Windows is not yet implemented yet.")
    else:
        raise ValueError(f"{sys.platform} is not implemented")


def document_to_pdf(docx_file: Path | str) -> None:
    """Create a PDF from a word document.

    Consider calling the relevant API's yourself
    if you need to add extra context to calls
    such as watermark arguments.
    """
    if isinstance(docx_file, str):
        docx_file = Path(docx_file)

    docx_file = docx_file.absolute()

    if sys.platform == "linux":
        _create_pdf_linux(docx_file)
    elif sys.platform == "win32":
        _create_pdf_windows(docx_file)
    elif sys.platform == "darwin":
        _create_pdf_macos(docx_file)
    else:
        raise ValueError(f"{sys.platform} is not implemented")
