"""Local package checks use isolated Git repositories and synthetic data."""

import hashlib
import importlib.util
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile


SCRIPT = Path(__file__).resolve().parents[1] / "build-windows-package.py"


class WindowsPackageTests(unittest.TestCase):
    def setUp(self):
        self.temporary = TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.base = Path(self.temporary.name)
        self.root = self.base / "project"
        self.root.mkdir()
        subprocess.run(["git", "init", "--quiet", str(self.root)], check=True)
        self.write("Dockerfile", "FROM python:3.12-slim\n")
        self.write(".dockerignore", ".local\n*.env\n*.xlsx\n")
        self.write("compose.windows.yml", "services: {}\n")
        self.write("docs/windows-local.md", "Windows deneme yönergesi.\n")
        self.write("README.md", "Vefeyât\n")
        self.write("scripts/build-windows-package.py", SCRIPT.read_text())
        self.output = self.base / "delivery.zip"

    def write(self, name, content):
        path = self.root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")
        return path

    def run_builder(self, *arguments):
        return subprocess.run(
            [sys.executable, str(self.root / "scripts/build-windows-package.py"), "--output", str(self.output), *arguments],
            cwd=self.base, capture_output=True, text=True,
        )

    def test_sources_are_packaged_without_hidden_credentials_or_runtime_data(self):
        self.write(".gitignore", "ignored.txt\n")
        self.write("ignored.txt", "Ignored output.\n")
        self.write("untracked.py", "print('development source')\n")
        self.write("vefeyat/settings.py", "DATABASE_PASSWORD = 'injected-at-runtime'\n")
        for name in (
            ".env", ".env.example", ".local/windows.env", ".local/dev-database.json", ".venv/config.py",
            "node_modules/dep/index.js", "config/passwords.txt", "config/private.key",
            "database.sqlite3", "database.sqlite3-wal", "db.json", "cached.xlsx", "nested/.config/settings.json",
            "nested/.dockerignore", "nested/.gitignore",
        ):
            self.write(name, "SYNTHETIC-PRIVATE-VALUE")
        # Even accidentally tracked sensitive files must never reach the ZIP.
        subprocess.run(["git", "-C", str(self.root), "add", "--force", "."], check=True)
        subprocess.run(["git", "-C", str(self.root), "rm", "--cached", "--quiet", "ignored.txt"], check=True)
        result = self.run_builder()
        self.assertEqual(result.returncode, 0, result.stderr)
        with ZipFile(self.output) as archive:
            names = set(archive.namelist())
            self.assertIn("vefeyat/untracked.py", names)
            self.assertIn("vefeyat/Dockerfile", names)
            self.assertIn("vefeyat/.dockerignore", names)
            self.assertIn("vefeyat/.gitignore", names)
            self.assertIn("vefeyat/vefeyat/settings.py", names)
            self.assertIn(b".local", archive.read("vefeyat/.dockerignore"))
            self.assertIn("vefeyat/compose.windows.yml", names)
            self.assertIn("vefeyat/docs/windows-local.md", names)
            self.assertNotIn("vefeyat/ignored.txt", names)
            self.assertFalse(any("/.local/" in name for name in names))
            self.assertNotIn("vefeyat/nested/.dockerignore", names)
            self.assertNotIn("vefeyat/nested/.gitignore", names)
            self.assertTrue(all(name.startswith("vefeyat/") for name in names))
            for name in names:
                self.assertNotIn(b"SYNTHETIC-PRIVATE-VALUE", archive.read(name))
            sums = archive.read("vefeyat/SHA256SUMS.txt").decode().splitlines()
            checksummed = set()
            for line in sums:
                digest, relative = line.split("  ", 1)
                member = "vefeyat/" + relative
                self.assertEqual(hashlib.sha256(archive.read(member)).hexdigest(), digest)
                checksummed.add(member)
            self.assertEqual(checksummed, names - {"vefeyat/SHA256SUMS.txt"})

    def test_explicit_workbook_is_the_only_runtime_data_exception(self):
        workbook = self.base / "private-name.xlsx"
        with ZipFile(workbook, "w") as archive:
            archive.writestr("[Content_Types].xml", "<Types/>")
            archive.writestr("xl/workbook.xml", "<workbook/>")
        result = self.run_builder("--source", str(workbook))
        self.assertEqual(result.returncode, 0, result.stderr)
        with ZipFile(self.output) as archive:
            self.assertEqual(archive.read("vefeyat/.local/source/people.xlsx"), workbook.read_bytes())
            self.assertEqual([name for name in archive.namelist() if "/.local/" in name], ["vefeyat/.local/source/people.xlsx"])
            self.assertFalse(any("private-name" in name for name in archive.namelist()))

    def test_missing_required_file_preserves_existing_output(self):
        self.output.write_bytes(b"previous successful package")
        for name in ("Dockerfile", ".dockerignore", "compose.windows.yml", "docs/windows-local.md"):
            with self.subTest(missing=name):
                path = self.root / name
                content = path.read_text()
                path.unlink()
                result = self.run_builder()
                self.assertNotEqual(result.returncode, 0)
                self.assertIn(name, result.stderr)
                self.assertEqual(self.output.read_bytes(), b"previous successful package")
                self.assertEqual(list(self.base.glob(".vefeyat-package-*")), [])
                self.write(name, content)

    def test_symlink_cannot_pull_an_external_file_into_the_package(self):
        external = self.base / "private.txt"
        external.write_text("Private external value")
        (self.root / "innocent.txt").symlink_to(external)
        result = self.run_builder()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("Sembolik", result.stderr)
        self.assertFalse(self.output.exists())

    def test_fake_xlsx_is_rejected_without_producing_a_zip(self):
        fake = self.base / "credentials.xlsx"
        fake.write_text("Not a workbook")
        result = self.run_builder("--source", str(fake))
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(self.output.exists())

    def test_symlink_workbook_and_output_paths_are_rejected(self):
        source = self.base / "real.xlsx"
        with ZipFile(source, "w") as archive:
            archive.writestr("[Content_Types].xml", "<Types/>")
            archive.writestr("xl/workbook.xml", "<workbook/>")
        linked_source = self.base / "linked.xlsx"
        linked_source.symlink_to(source)
        result = self.run_builder("--source", str(linked_source))
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(self.output.exists())
        protected = self.base / "existing.zip"
        protected.write_bytes(b"preserve this file")
        self.output.symlink_to(protected)
        result = self.run_builder()
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(protected.read_bytes(), b"preserve this file")

    def test_windows_case_collision_is_rejected(self):
        self.write("notes/Guide.md", "First file")
        self.write("notes/guide.md", "Second file")
        result = self.run_builder()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("çakışan", result.stderr)
        self.assertFalse(self.output.exists())

    def test_unsafe_archive_paths_are_rejected(self):
        spec = importlib.util.spec_from_file_location("windows_package", SCRIPT)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        for name in ("/etc/passwd", "../outside", "a/../../escape", "a\\..\\escape", "C:/outside", "a//b", "a/./b", "CON.txt", "file. ", "line\nbreak"):
            with self.subTest(name=name):
                with self.assertRaises(module.PackageError):
                    module.validate_relative_path(name)


if __name__ == "__main__":
    unittest.main()
