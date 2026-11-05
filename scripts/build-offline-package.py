#!/usr/bin/env python3
"""Build the server-free Windows Word add-in package using the standard library.

An optional private CSV is copied directly into the ZIP, never into the repository.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
from pathlib import Path
import tempfile
import zipfile

ROOT = Path(__file__).resolve().parents[1]
DESKTOP = ROOT / "desktop"
ARCHIVE_ROOT = "Vefeyat-Windows/"
HEADERS = ("Kısa Ad", "Tam Ad", "Orijinal Ad", "Vefat Tarihi", "Tanıtım", "Kaynak Link")
SOURCES = (
    "Kur.cmd", "Kaldir.cmd", "Install.ps1", "Uninstall.ps1",
    "Core.cs", "CatalogForm.cs", "Connect.cs", "SelectionRules.cs", "KULLANIM.md",
)
ZIP_DATE = (2026, 1, 1, 0, 0, 0)


def validate_csv(data: bytes) -> int:
    """Validate UTF-8 CSV without transforming any of its original field values."""
    try:
        content = data.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("CSV dosyası UTF-8 kodlamasında olmalıdır.") from exc
    if "\x00" in content:
        raise ValueError("CSV dosyası NUL karakteri içeremez.")
    reader = csv.reader(io.StringIO(content, newline=""), delimiter=";", strict=True)
    try:
        headers = next(reader, None)
        if headers != list(HEADERS):
            raise ValueError("CSV başlıkları sırasıyla şu altı alan olmalıdır: " + ";".join(HEADERS))
        count = 0
        for row in reader:
            if not row or all(not field.strip() for field in row):
                continue
            if len(row) != len(HEADERS):
                raise ValueError(f"CSV satırı {reader.line_num}: altı sütun bekleniyor.")
            if not row[0].strip():
                raise ValueError(f"CSV satırı {reader.line_num}: Kısa Ad boş olamaz.")
            count += 1
    except csv.Error as exc:
        raise ValueError(f"CSV okunamadı: {exc}") from exc
    if count == 0:
        raise ValueError("CSV en az bir şahıs kaydı içermelidir.")
    return count


def package_files(csv_path: Path | None = None) -> dict[str, bytes]:
    selected_csv = csv_path if csv_path is not None else DESKTOP / "data" / "people.csv"
    seed = selected_csv.read_bytes()
    row_count = validate_csv(seed)
    contents = {name: (DESKTOP / name).read_bytes() for name in SOURCES}
    contents["data/people.csv"] = seed
    metadata = {
        "format": 1,
        "product": "Vefeyât Windows Word",
        "seed_records": row_count,
        "seed_sha256": hashlib.sha256(seed).hexdigest(),
        "sample_data": csv_path is None,
        "requires_network": False,
        "requires_server": False,
        "build": "Windows üzerinde yerleşik .NET Framework C# derleyicisi ile çevrimdışı",
    }
    contents["paket-bilgisi.json"] = (json.dumps(metadata, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
    sums = "".join(f"{hashlib.sha256(contents[name]).hexdigest()}  {name}\n" for name in sorted(contents))
    contents["SHA256SUMS.txt"] = sums.encode("utf-8")
    return contents


def build_package(output: Path, csv_path: Path | None = None) -> dict[str, object]:
    contents = package_files(csv_path)
    output = output.resolve()
    protected_sources = {(DESKTOP / name).resolve() for name in SOURCES}
    protected_sources.add((DESKTOP / "data" / "people.csv").resolve())
    if csv_path is not None:
        protected_sources.add(csv_path.resolve())
    if output in protected_sources:
        raise ValueError("Çıktı ZIP dosyası bir kaynak veya CSV dosyasının üzerine yazılamaz.")
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(prefix=".vefeyat-", suffix=".zip", dir=output.parent, delete=False) as temporary:
            temporary_path = Path(temporary.name)
        with zipfile.ZipFile(temporary_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
            for name, data in sorted(contents.items()):
                info = zipfile.ZipInfo(ARCHIVE_ROOT + name, date_time=ZIP_DATE)
                info.compress_type = zipfile.ZIP_DEFLATED
                info.create_system = 0
                info.external_attr = 0o100644 << 16
                archive.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
        temporary_path.replace(output)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
    metadata = json.loads(contents["paket-bilgisi.json"])
    return {"output": str(output), "seed_records": metadata["seed_records"], "sha256": hashlib.sha256(output.read_bytes()).hexdigest()}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--csv", type=Path, help="Paket içindeki başlangıç verisi; UTF-8, noktalı virgülle ayrılmış altı sütunlu CSV")
    parser.add_argument("--output", required=True, type=Path, help="Oluşturulacak ZIP dosyası")
    args = parser.parse_args()
    try:
        result = build_package(args.output, args.csv)
    except (OSError, ValueError) as exc:
        parser.exit(1, f"Paket oluşturulamadı: {exc}\n")
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
