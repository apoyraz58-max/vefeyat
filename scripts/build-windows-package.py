#!/usr/bin/env python3
"""Build a local Windows trial ZIP without runtime state or credentials.

Run from any directory. Only Git-listed project sources are considered. An
explicit --source workbook is the sole permitted runtime-data exception.
"""

import argparse
import hashlib
import os
from pathlib import Path, PurePosixPath
import re
import stat
import subprocess
import sys
import tempfile
import unicodedata
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile, ZipInfo


REQUIRED_FILES = ("Dockerfile", ".dockerignore", "compose.windows.yml", "docs/windows-local.md")
PUBLIC_ROOT_DOTFILES = {".dockerignore", ".gitignore"}
SOURCE_ARCHIVE_PATH = "vefeyat/.local/source/people.xlsx"
MANIFEST_ARCHIVE_PATH = "vefeyat/SHA256SUMS.txt"
EXCLUDED_DIRECTORIES = {
    "node_modules", "venv", "env", "__pycache__", "pgdata", "secrets",
    "credentials", "certs", "certificates", "uploads", "media", "staticfiles",
    "htmlcov", "playwright-report", "test-results",
}
EXCLUDED_SUFFIXES = {
    ".env", ".pem", ".key", ".p12", ".pfx", ".crt", ".db", ".sqlite",
    ".sqlite3", ".sql", ".dump", ".bak", ".log", ".pyc", ".pyo",
    ".xlsx", ".xls", ".csv", ".jsonl", ".zip", ".tar", ".gz", ".tgz",
    ".7z", ".rar",
}
SENSITIVE_NAME = re.compile(
    r"(?:^|[_.-])(?:secrets?|credentials?|passwords?|passwd|tokens?)(?:$|[_.-])",
    re.IGNORECASE,
)
WINDOWS_DEVICES = {"con", "prn", "aux", "nul", *(f"com{i}" for i in range(1, 10)), *(f"lpt{i}" for i in range(1, 10))}


class PackageError(Exception):
    pass


def validate_relative_path(name):
    """Accept unambiguous relative paths that Windows can safely extract."""
    if not name or name.startswith("/") or "\\" in name or ":" in name:
        raise PackageError(f"Güvensiz paket yolu: {name!r}")
    parts = name.split("/")
    if any(part in {"", ".", ".."} for part in parts):
        raise PackageError(f"Güvensiz paket yolu: {name!r}")
    for part in parts:
        if any(ord(char) < 32 or ord(char) == 127 for char in part):
            raise PackageError("Paket yolunda kontrol karakteri bulunamaz.")
        if part.endswith((" ", ".")) or any(char in part for char in '<>"|?*'):
            raise PackageError(f"Windows için geçersiz paket yolu: {name!r}")
        if part.split(".", 1)[0].casefold() in WINDOWS_DEVICES:
            raise PackageError(f"Windows aygıt adı paketlenemez: {name!r}")
    return PurePosixPath(name)


def excluded_source(relative):
    parts = relative.parts
    # These repository-root rules contain no runtime values and keep local
    # source data and credentials out of Docker images and future Git commits.
    if len(parts) == 1 and parts[0] in PUBLIC_ROOT_DOTFILES:
        return False
    if any(part.startswith(".") or part.casefold() in EXCLUDED_DIRECTORIES for part in parts):
        return True
    filename = parts[-1].casefold()
    return (
        Path(filename).suffix in EXCLUDED_SUFFIXES
        or ".sqlite" in filename
        or bool(SENSITIVE_NAME.search(filename))
        or filename.startswith(("database.", "db."))
        or filename in {"dev-database.json", "pgpass", "pg_hba.conf", "pg_ident.conf", "postgresql.conf", "id_rsa", "id_ed25519"}
    )


def regular_file(path, *, within=None):
    """Reject links in both the file and its parent directories."""
    absolute = path.absolute()
    for component in (absolute, *absolute.parents):
        if component.is_symlink():
            raise PackageError(f"Sembolik bağlantı paketlenemez: {path}")
    if not absolute.is_file():
        raise PackageError(f"Normal dosya bulunamadı: {path}")
    resolved = absolute.resolve()
    if within is not None and not resolved.is_relative_to(within):
        raise PackageError(f"Depo dışına çıkan kaynak yolu: {path}")
    return resolved


def git_sources(root):
    result = subprocess.run(
        ["git", "-C", str(root), "ls-files", "--cached", "--others", "--exclude-standard", "-z"],
        check=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
    )
    if result.returncode:
        raise PackageError("Git kaynak listesi okunamadı; betiği Git deposunda kullanın.")
    try:
        return sorted(set(item.decode("utf-8") for item in result.stdout.split(b"\0") if item))
    except UnicodeDecodeError as error:
        raise PackageError("Kaynak dosya adları UTF-8 olmalıdır.") from error


def validate_workbook(path):
    source = regular_file(path)
    if source.suffix.casefold() != ".xlsx":
        raise PackageError("--source gerçek bir .xlsx çalışma kitabı olmalıdır.")
    try:
        with ZipFile(source) as book:
            if not {"[Content_Types].xml", "xl/workbook.xml"}.issubset(book.namelist()) or book.testzip():
                raise PackageError("--source geçerli bir XLSX çalışma kitabı değil.")
    except (BadZipFile, OSError) as error:
        raise PackageError("--source geçerli bir XLSX çalışma kitabı değil.") from error
    return source


def zip_info(name):
    info = ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
    info.compress_type = ZIP_DEFLATED
    info.create_system = 3
    info.external_attr = (stat.S_IFREG | 0o644) << 16
    return info


def verify_archive(path, expected_names):
    with ZipFile(path) as archive:
        if archive.testzip():
            raise PackageError("ZIP bütünlük denetimi başarısız oldu.")
        names = archive.namelist()
        if len(names) != len(set(names)) or set(names) != expected_names:
            raise PackageError("ZIP içeriği beklenen kaynak listesiyle uyuşmuyor.")
        for info in archive.infolist():
            relative = validate_relative_path(info.filename)
            if relative.parts[0] != "vefeyat" or len(relative.parts) < 2:
                raise PackageError("ZIP dosyaları vefeyat/ altında olmalıdır.")
            if stat.S_ISLNK(info.external_attr >> 16):
                raise PackageError("ZIP sembolik bağlantı içeriyor.")
            if info.filename != SOURCE_ARCHIVE_PATH and excluded_source(PurePosixPath(*relative.parts[1:])):
                raise PackageError("ZIP yasaklı yapılandırma veya veri dosyası içeriyor.")


def build_package(root, output, source=None):
    root = root.resolve()
    output = output.absolute()
    if output.suffix.casefold() != ".zip":
        raise PackageError("--output yolu .zip uzantılı olmalıdır.")
    if output.is_symlink() or any(parent.is_symlink() for parent in output.parents):
        raise PackageError("Çıktı yolu sembolik bağlantı içeremez.")
    files = {}
    windows_names = set()
    for name in git_sources(root):
        relative = validate_relative_path(name)
        if excluded_source(relative) or root.joinpath(*relative.parts).absolute() == output:
            continue
        archive_name = f"vefeyat/{relative.as_posix()}"
        if unicodedata.normalize("NFC", archive_name).casefold() == MANIFEST_ARCHIVE_PATH.casefold():
            raise PackageError("SHA256SUMS.txt paket oluşturucusu için ayrılmıştır.")
        normalized = unicodedata.normalize("NFC", archive_name).casefold()
        if normalized in windows_names:
            raise PackageError(f"Windows'ta çakışan dosya adı: {name}")
        windows_names.add(normalized)
        files[archive_name] = regular_file(root.joinpath(*relative.parts), within=root)
    missing = [name for name in REQUIRED_FILES if f"vefeyat/{name}" not in files]
    if missing:
        raise PackageError("Windows paketi hazır değil; gerekli dosyalar eksik: " + ", ".join(missing))
    if source is not None:
        files[SOURCE_ARCHIVE_PATH] = validate_workbook(source)

    # Never leave a partial ZIP or replace a previous package after a failure.
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(prefix=".vefeyat-package-", suffix=".tmp", dir=output.parent, delete=False) as stream:
            temporary = Path(stream.name)
        checksums = []
        with ZipFile(temporary, "w", compression=ZIP_DEFLATED) as archive:
            for name, path in sorted(files.items()):
                regular_file(path, within=None if name == SOURCE_ARCHIVE_PATH else root)
                digest = hashlib.sha256()
                flags = os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0)
                descriptor = os.open(path, flags)
                with os.fdopen(descriptor, "rb") as incoming, archive.open(zip_info(name), "w") as outgoing:
                    if not stat.S_ISREG(os.fstat(incoming.fileno()).st_mode):
                        raise PackageError(f"Normal dosya paketlenemedi: {path}")
                    while chunk := incoming.read(1024 * 1024):
                        digest.update(chunk)
                        outgoing.write(chunk)
                checksums.append(f"{digest.hexdigest()}  {name.removeprefix('vefeyat/')}\n")
            archive.writestr(zip_info(MANIFEST_ARCHIVE_PATH), "".join(checksums).encode("utf-8"))
        verify_archive(temporary, set(files) | {MANIFEST_ARCHIVE_PATH})
        os.replace(temporary, output)
        temporary = None
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
    return len(files)


def main(argv=None):
    parser = argparse.ArgumentParser(description="Vefeyât Windows deneme ZIP'i oluşturur; ağa veya yayına çıkmaz.")
    parser.add_argument("--output", type=Path, required=True, help="Oluşturulacak .zip dosyası")
    parser.add_argument("--source", type=Path, help="İsteğe bağlı, kullanıcıya ait XLSX verisi")
    arguments = parser.parse_args(argv)
    try:
        count = build_package(Path(__file__).resolve().parents[1], arguments.output, arguments.source)
    except (PackageError, OSError, subprocess.SubprocessError) as error:
        print(f"Paket oluşturulamadı: {error}", file=sys.stderr)
        return 1
    print(f"Windows deneme paketi oluşturuldu: {arguments.output} ({count} dosya + SHA256SUMS.txt)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
