"""Behavior checks for the offline distribution and user CSV boundary."""
import csv
import hashlib
import importlib.util
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile

SCRIPT = Path(__file__).resolve().parents[1] / "build-offline-package.py"
SPEC = importlib.util.spec_from_file_location("build_offline_package", SCRIPT)
PACKAGE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(PACKAGE)


def seed_bytes(rows=None):
    content = io.StringIO(newline="")
    writer = csv.writer(content, delimiter=";", lineterminator="\r\n")
    writer.writerow(PACKAGE.HEADERS)
    writer.writerows(rows or [["Özel Şahıs", "İsim; unvan", "عربي", "(ö. 1/2)", 'Birinci satır\nİkinci "satır".', "https://example.invalid/kaynak"]])
    return b"\xef\xbb\xbf" + content.getvalue().encode("utf-8")


class CsvValidationTests(unittest.TestCase):
    def test_preserves_quotes_multiline_turkish_arabic_and_bom(self):
        self.assertEqual(PACKAGE.validate_csv(seed_bytes()), 1)

    def test_rejects_wrong_headers_wrong_column_count_and_empty_name(self):
        bad_values = [
            b"Name;Date\nAli;1000\n",
            (";".join(PACKAGE.HEADERS) + "\nAli;1000\n").encode("utf-8"),
            seed_bytes([["", "", "", "", "", "URL"]]),
            (";".join(PACKAGE.HEADERS) + "\n").encode("utf-8"),
            b"\xff",
            seed_bytes().replace(b"URL", b"\x00") + b"\x00",
        ]
        for value in bad_values:
            with self.subTest(value=value[:40]), self.assertRaises(ValueError):
                PACKAGE.validate_csv(value)


class PackageTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.desktop = self.root / "desktop"
        (self.desktop / "data").mkdir(parents=True)
        for name in PACKAGE.SOURCES:
            (self.desktop / name).write_text("allowed source: " + name, encoding="utf-8")
        self.sample = seed_bytes([["[Örnek] Hayalî", "---", "---", "---", "Hayalî örnek", ""]])
        (self.desktop / "data" / "people.csv").write_bytes(self.sample)
        # These unrelated assets must never enter an end-user package.
        (self.desktop / "secret.env").write_text("private=true", encoding="utf-8")
        (self.desktop / "Dockerfile").write_text("FROM unrelated", encoding="utf-8")
        self.patcher = patch.object(PACKAGE, "DESKTOP", self.desktop)
        self.patcher.start()
        self.addCleanup(self.patcher.stop)

    def test_deterministic_allowlist_and_integrity_manifest(self):
        first, second = self.root / "first.zip", self.root / "second.zip"
        PACKAGE.build_package(first)
        PACKAGE.build_package(second)
        self.assertEqual(first.read_bytes(), second.read_bytes())
        with zipfile.ZipFile(first) as archive:
            expected = set(PACKAGE.SOURCES) | {"data/people.csv", "paket-bilgisi.json", "SHA256SUMS.txt"}
            self.assertEqual(set(archive.namelist()), {PACKAGE.ARCHIVE_ROOT + name for name in expected})
            metadata = json.loads(archive.read(PACKAGE.ARCHIVE_ROOT + "paket-bilgisi.json"))
            self.assertTrue(metadata["sample_data"])
            self.assertEqual(metadata["seed_records"], 1)
            for line in archive.read(PACKAGE.ARCHIVE_ROOT + "SHA256SUMS.txt").decode("utf-8").splitlines():
                digest, name = line.split("  ", 1)
                self.assertEqual(hashlib.sha256(archive.read(PACKAGE.ARCHIVE_ROOT + name)).hexdigest(), digest)

    def test_private_csv_goes_into_archive_without_overwriting_sample(self):
        private_csv = self.root / "şahıslar.csv"
        supplied = seed_bytes()
        private_csv.write_bytes(supplied)
        output = self.root / "private.zip"
        result = PACKAGE.build_package(output, private_csv)
        self.assertEqual(result["seed_records"], 1)
        self.assertEqual((self.desktop / "data" / "people.csv").read_bytes(), self.sample)
        self.assertEqual(private_csv.read_bytes(), supplied)
        with zipfile.ZipFile(output) as archive:
            self.assertEqual(archive.read(PACKAGE.ARCHIVE_ROOT + "data/people.csv"), supplied)
            metadata = json.loads(archive.read(PACKAGE.ARCHIVE_ROOT + "paket-bilgisi.json"))
            self.assertFalse(metadata["sample_data"])
            self.assertNotIn("şahıslar.csv", "\n".join(archive.namelist()))

    def test_failed_validation_preserves_existing_output(self):
        output = self.root / "existing.zip"
        output.write_bytes(b"previous artifact")
        invalid = self.root / "invalid.csv"
        invalid.write_bytes(b"invalid")
        with self.assertRaises(ValueError):
            PACKAGE.build_package(output, invalid)
        self.assertEqual(output.read_bytes(), b"previous artifact")

    def test_cannot_replace_source_or_input_csv(self):
        for path in (self.desktop / "Core.cs", self.desktop / "data" / "people.csv"):
            before = path.read_bytes()
            with self.subTest(path=path), self.assertRaises(ValueError):
                PACKAGE.build_package(path)
            self.assertEqual(path.read_bytes(), before)
        private_csv = self.root / "input.csv"
        private_csv.write_bytes(seed_bytes())
        with self.assertRaises(ValueError):
            PACKAGE.build_package(private_csv, private_csv)


if __name__ == "__main__":
    unittest.main()
