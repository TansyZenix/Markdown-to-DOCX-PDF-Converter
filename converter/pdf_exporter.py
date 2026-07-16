from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path


class LibreOfficeNotFoundError(RuntimeError):
    pass


def find_libreoffice(executable: str | None = None) -> str | None:
    if executable:
        return executable if Path(executable).is_file() else None
    for candidate in (shutil.which("soffice"), shutil.which("soffice.exe"), r"C:\\Program Files\\LibreOffice\\program\\soffice.exe"):
        if candidate and Path(candidate).is_file():
            return candidate
    return None


def find_word() -> str | None:
    for candidate in (shutil.which("WINWORD.EXE"), r"C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE", r"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\WINWORD.EXE"):
        if candidate and Path(candidate).is_file():
            return candidate
    return None


def build_pdf_command(executable: Path, docx_path: Path, output_directory: Path, profile_directory: Path | None = None) -> list[str]:
    command = [str(executable), "--headless"]
    if profile_directory:
        command.append(f"-env:UserInstallation={profile_directory.resolve().as_uri()}")
    return command + ["--convert-to", "pdf", "--outdir", str(output_directory), str(docx_path)]


def _export_with_libreoffice(executable: str, docx_path: Path, output_directory: Path) -> Path:
    with tempfile.TemporaryDirectory(prefix="md2docx-lo-") as profile:
        result = subprocess.run(build_pdf_command(Path(executable), docx_path, output_directory, Path(profile)), capture_output=True, text=True, check=False)
    pdf_path = output_directory / f"{docx_path.stem}.pdf"
    if result.returncode != 0 or not pdf_path.is_file() or pdf_path.stat().st_size == 0:
        raise RuntimeError(f"LibreOffice PDF export failed: {result.stderr.strip() or result.stdout.strip()}")
    return pdf_path


def _export_with_word(docx_path: Path, output_directory: Path) -> Path:
    try:
        from win32com.client import DispatchEx
    except ImportError as error:
        raise LibreOfficeNotFoundError("LibreOffice is unavailable and pywin32 is not installed for the Microsoft Word fallback.") from error
    pdf_path = output_directory / f"{docx_path.stem}.pdf"
    word = DispatchEx("Word.Application")
    word.Visible = False
    try:
        document = word.Documents.Open(str(docx_path), ReadOnly=True)
        document.SaveAs(str(pdf_path), FileFormat=17)
        document.Close(False)
    finally:
        word.Quit()
    if not pdf_path.is_file() or pdf_path.stat().st_size == 0:
        raise RuntimeError("Microsoft Word PDF export did not create a PDF file.")
    return pdf_path


def export_pdf(docx_path: Path, output_directory: Path, executable: str | None = None) -> Path:
    docx_path, output_directory = Path(docx_path).resolve(), Path(output_directory).resolve()
    output_directory.mkdir(parents=True, exist_ok=True)
    if executable:
        resolved = find_libreoffice(executable)
        if not resolved:
            raise LibreOfficeNotFoundError("LibreOffice was not found. Install it or pass --libreoffice C:\\path\\to\\soffice.exe")
        return _export_with_libreoffice(resolved, docx_path, output_directory)
    resolved = find_libreoffice()
    if resolved:
        return _export_with_libreoffice(resolved, docx_path, output_directory)
    if find_word():
        return _export_with_word(docx_path, output_directory)
    raise LibreOfficeNotFoundError("No PDF converter found. Install LibreOffice or Microsoft Word, then retry.")
