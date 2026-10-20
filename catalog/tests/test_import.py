"""Exercise the real Excel reader without depending on the private source file."""

from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory

from django.core.management import call_command
from django.core.management.base import CommandError
from django.test import TestCase
from openpyxl import Workbook

from catalog.models import ImportBatch, Person


HEADERS = ["Kısa Ad", "Tam Ad", "Orijinal Ad", "Vefat Tarihi", "Tanıtım", "Kaynak Link"]
ROWS = [
    ["İbn Rüşd", "Ebü'l-Velîd İbn Rüşd", "ابن رشد", "(ö. 520/1126)", "Fakih.", "https://example.org/ibn-rusd-dede"],
    ["İbn Rüşd", "---", "---", "(ö. 595/1198)", "Filozof ve hekim.", "https://example.org/ibn-rusd-torun"],
    ["Muhammed el-Mehdî", None, "محمد المهدي", "(gaybeti 260/874)", "Gaybet tarihi bildirilen şahıs.", "https://example.org/mehdi"],
]


class ImportPeopleTests(TestCase):
    def setUp(self):
        self.temporary = TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)

    def workbook(self, rows=None, name="people.xlsx"):
        path = Path(self.temporary.name) / name
        book = Workbook()
        sheet = book.active
        sheet.title = "alimler_251025"
        sheet.append(HEADERS)
        for row in rows if rows is not None else ROWS:
            sheet.append(row)
        book.save(path)
        book.close()
        return path

    def run_import(self, path):
        output = StringIO()
        call_command("import_people", str(path), stdout=output)
        return output.getvalue()

    def test_import_keeps_distinct_people_raw_dates_and_source_cells(self):
        self.run_import(self.workbook())

        self.assertEqual(Person.objects.count(), 3)
        self.assertEqual(Person.objects.filter(short_name="İbn Rüşd").count(), 2)
        batch = ImportBatch.objects.get()
        self.assertEqual(batch.row_count, 3)
        for row_number, original in enumerate(ROWS, 2):
            with self.subTest(row=row_number):
                person = Person.objects.get(source_url=original[-1])
                self.assertEqual(person.date_text, original[3])
                self.assertEqual(person.original_data, dict(zip(HEADERS, original)))
                self.assertEqual(person.source_sheet, "alimler_251025")
                self.assertEqual(person.source_row, row_number)
                self.assertEqual(person.import_batch_id, batch.pk)
        missing = Person.objects.get(source_url=ROWS[1][-1])
        self.assertEqual(missing.full_name, "")
        self.assertEqual(missing.original_name, "")
        self.assertEqual(missing.original_data["Tam Ad"], "---")

    def test_period_expression_is_not_reinterpreted_as_a_death_year(self):
        rows = [list(row) for row in ROWS]
        rows[2][3] = "(III./IX. yüzyılın ikinci yarısı)"
        self.run_import(self.workbook(rows))
        person = Person.objects.get(source_url=rows[2][-1])
        self.assertEqual(person.date_text, rows[2][3])
        self.assertEqual(person.original_data["Vefat Tarihi"], rows[2][3])

    def test_reimport_same_file_preserves_editor_changes(self):
        path = self.workbook()
        self.run_import(path)
        person = Person.objects.get(source_url=ROWS[0][-1])
        person.short_name = "İbn Rüşd el-Cedd"
        person.description = "Editör tarafından doğrulanan açıklama."
        person.save()
        original_id = person.pk

        self.run_import(path)

        person.refresh_from_db()
        self.assertEqual(Person.objects.count(), 3)
        self.assertEqual(ImportBatch.objects.count(), 1)
        self.assertEqual(person.pk, original_id)
        self.assertEqual(person.short_name, "İbn Rüşd el-Cedd")
        self.assertEqual(person.description, "Editör tarafından doğrulanan açıklama.")
        self.assertEqual(person.original_data["Kısa Ad"], "İbn Rüşd")

    def test_invalid_later_row_leaves_no_partial_import(self):
        for invalid in ("http://example.org/insecure", "javascript:alert(1)", "not a URL"):
            with self.subTest(source_url=invalid):
                rows = [list(row) for row in ROWS]
                rows[-1][-1] = invalid
                with self.assertRaises(CommandError):
                    self.run_import(self.workbook(rows))
                self.assertFalse(Person.objects.exists())
                self.assertFalse(ImportBatch.objects.exists())

    def test_duplicate_source_in_workbook_rolls_back_all_rows(self):
        rows = [list(row) for row in ROWS]
        rows[-1][-1] = rows[0][-1]
        with self.assertRaises(CommandError):
            self.run_import(self.workbook(rows))
        self.assertFalse(Person.objects.exists())
        self.assertFalse(ImportBatch.objects.exists())

    def test_source_conflict_with_existing_person_preserves_database(self):
        existing = Person.objects.create(
            short_name="Editör kaydı", date_text="(tarih bilinmiyor)",
            source_url=ROWS[-1][-1], description="Mevcut açıklama.",
        )
        with self.assertRaises(CommandError):
            self.run_import(self.workbook())
        self.assertEqual(list(Person.objects.values_list("pk", flat=True)), [existing.pk])
        existing.refresh_from_db()
        self.assertEqual(existing.short_name, "Editör kaydı")
        self.assertEqual(existing.description, "Mevcut açıklama.")
        self.assertFalse(ImportBatch.objects.exists())

    def test_excel_formula_is_not_imported_as_a_name(self):
        rows = [list(row) for row in ROWS]
        rows[-1][0] = '=HYPERLINK("https://example.org/", "Başlık")'
        with self.assertRaises(CommandError):
            self.run_import(self.workbook(rows))
        self.assertFalse(Person.objects.exists())
        self.assertFalse(ImportBatch.objects.exists())
