import hashlib
import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
import warnings
from collections.abc import Callable
from pathlib import Path

log = logging.getLogger(__name__)


def get_sha1(path: Path) -> str | None:
    """
    Get the SHA1 checksum of the file at `path`.
    """
    if not path.exists():
        return None

    sha1sum = hashlib.sha1()
    with open(path, "rb") as src:
        block = src.read(2 ** 16)
        while len(block) != 0:
            sha1sum.update(block)
            block = src.read(2 ** 16)
    return sha1sum.hexdigest().lower()


def _create_pdf_windows(docx_file: Path) -> None:
    import win32com.client

    word = win32com.client.Dispatch("Word.Application")
    wdFormatPDF = 17

    docx_filepath = docx_file
    pdf_filepath = Path(f"{docx_file.stem}.pdf").absolute().resolve()
    doc = word.Documents.Open(str(docx_filepath))
    try:
        doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
    except:
        raise
    finally:
        doc.Close(0)

    word.Quit()


def _create_pdf_linux(docx_file: Path) -> None:
    try:
        subprocess.call(
            [
                "libreoffice",
                "--convert-to",
                "pdf",
                str(docx_file),
            ],
            timeout=5,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except subprocess.TimeoutExpired:
        # New versions of LibreOffice appear to
        # hang even after the PDF are created hence
        # why we enforce a process timeout
        log.debug(
            "DOCX to PDF call timed out after 5 seconds. "
            "This is likely fine, but if not please open an issue."
        )


def _create_pdf_macos(docx_file: Path) -> None:
    log.warning("DOCX -> PDF on mac is untested. Any issues please raise an issue.")
    script = (Path(__file__).parent / "convert.jxa").absolute().resolve()
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


def export_libre_macro(macro_folder: Path | None = None) -> None:
    """
    Automatically moves the LibreOffice macro file to `macro_folder`.

    Warning, this overrides Module1.xba

    :py:class:`Path` is where your macros live (leave None to let package choose location)
    """
    if macro_folder is None:
        platform_paths = {
            "win32": Path(os.path.expandvars("%APPDATA%"), "LibreOffice/4/user/basic/Standard"),
            "linux": Path("~/.config/libreoffice/4/user/basic/Standard").expanduser(),
            "darwin": Path("~/Library/Application Support/LibreOffice/4/user/basic/Standard").expanduser()
        }

        try:
            macro_folder = platform_paths[sys.platform]
        except KeyError as e:
            raise ValueError(f"Unsupported platform: {sys.platform}") from e

    expect_macro_sha1 = "539afdb97c8fb21a0cd08143d6a531d7d683df21"

    target_macro_path = macro_folder / "Module1.xba"
    target_macro_sha1 = get_sha1(target_macro_path)

    if expect_macro_sha1 == target_macro_sha1:
        return  # No changes required

    stored_macro_path = Path(__file__).parent / "macros/Module1.xba"
    stored_macro_sha1 = get_sha1(stored_macro_path)

    if expect_macro_sha1 != stored_macro_sha1:
        raise ValueError(
            f"Unexpected SHA1 checksum for stored macro: {stored_macro_path.name}"
            f" (expected={expect_macro_sha1}, actual={stored_macro_sha1}"
        )

    log.info(f"Overwriting macro at location {target_macro_path}")
    shutil.copy(stored_macro_path, target_macro_path)


def update_toc(docx_file: Path | str) -> None:
    """
    Update the table of contents and indexes within a Word document.

    https://github.com/python-openxml/python-docx/issues/1207#issuecomment-1924053420
    """
    docx_file = Path(docx_file).absolute().resolve()
    callback: Callable[[Path], ...]

    if sys.platform == "linux":
        callback = _update_toc_linux
    elif sys.platform == "win32":
        callback = _update_toc_windows
    elif sys.platform == "darwin":
        callback = _update_toc_macos
    else:
        raise ValueError(f"Unsupported platform: {sys.platform}")

    with tempfile.TemporaryDirectory() as temp_dir:  # https://stackoverflow.com/questions/23212435
        temp_path = Path(temp_dir, "temp.docx")

        shutil.copy(docx_file, temp_path)
        callback(temp_path)
        shutil.copy(temp_path, docx_file)


def _update_toc_linux(docx_file: Path) -> None:
    """
    Helper method for Linux (Call UpdateTOC binding on filepath))
    """
    subprocess.call(
        [
            "libreoffice",
            "--headless",
            f"macro:///Standard.Module1.UpdateTOC({docx_file})",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def _update_toc_windows(docx_file: Path) -> None:
    """
    Helper method for Windows (Call UpdateTOC binding on filepath)
    """
    subprocess.call(
        [
            "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
            f"macro:///Standard.Module1.UpdateTOC({docx_file})",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def _update_toc_macos(docx_file: Path) -> None:
    """
    Helper method for macOS (Call UpdateTOC binding on filepath)
    """
    subprocess.call(
        [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "--headless",
            f"macro:///Standard.Module1.UpdateTOC({docx_file})",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


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
