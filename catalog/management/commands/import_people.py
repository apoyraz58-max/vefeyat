import hashlib
from pathlib import Path

from django.core.management.base import BaseCommand, CommandError
from django.core.validators import URLValidator
from django.core.exceptions import ValidationError
from django.db import transaction
from openpyxl import load_workbook

from catalog.models import ImportBatch, Person
from catalog.search import normalize_name

HEADERS = ["Kısa Ad", "Tam Ad", "Orijinal Ad", "Vefat Tarihi", "Tanıtım", "Kaynak Link"]
MISSING = {"", "---", "--", "-", "—", "–"}


class Command(BaseCommand):
    help = "Excel şahıs verisini kayıpsız ve tekrar çalıştırılabilir biçimde aktarır."

    def add_arguments(self, parser):
        parser.add_argument("file", type=Path)
        parser.add_argument("--sheet", default="alimler_251025")

    def handle(self, *args, **options):
        path = options["file"]
        if not path.is_file():
            raise CommandError("Excel dosyası bulunamadı.")
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        if ImportBatch.objects.filter(sha256=digest).exists():
            self.stdout.write("Bu dosya daha önce aktarıldı; mevcut kayıtlar değiştirilmedi.")
            return
        try:
            workbook = load_workbook(path, read_only=True, data_only=False)
            if options["sheet"] not in workbook.sheetnames:
                raise CommandError("İstenen çalışma sayfası bulunamadı.")
            sheet = workbook[options["sheet"]]
            iterator = sheet.iter_rows(values_only=True)
            if list(next(iterator, ())) != HEADERS:
                raise CommandError("Excel başlıkları beklenen altı sütunla aynı değil.")
            records = []
            seen_urls = set()
            for row_number, row in enumerate(iterator, 2):
                if not any(v is not None for v in row):
                    continue
                if any(not isinstance(v, (str, type(None))) or (isinstance(v, str) and v.startswith("=")) for v in row):
                    raise CommandError(f"{row_number}. satır: metin dışında değer veya formül var.")
                cleaned = ["" if value is None or value.strip() in MISSING else value for value in row]
                short, full, original, date, description, url = cleaned
                try:
                    URLValidator(schemes=["https"])(url)
                except ValidationError:
                    raise CommandError(f"{row_number}. satır: HTTPS kaynak URL'si geçersiz.")
                if not short.strip() or not date.strip() or not description.strip():
                    raise CommandError(f"{row_number}. satır: kısa ad, tarih ifadesi ve tanıtım gereklidir.")
                if url in seen_urls:
                    raise CommandError(f"{row_number}. satır: kaynak URL'si tekrarlanıyor; otomatik birleştirme yapılmadı.")
                if len(short) > 500 or len(date) > 500 or len(url) > 1000:
                    raise CommandError(f"{row_number}. satır: alan uzunluğu sınırı aşılıyor.")
                seen_urls.add(url)
                records.append(Person(
                    short_name=short, full_name=full, original_name=original, date_text=date,
                    description=description, source_url=url, short_norm=normalize_name(short),
                    full_norm=normalize_name(full), original_norm=normalize_name(original),
                    original_data=dict(zip(HEADERS, row)), source_sheet=sheet.title, source_row=row_number,
                ))
        finally:
            if "workbook" in locals():
                workbook.close()
        if not records:
            raise CommandError("Dosyada aktarılacak kayıt yok.")
        with transaction.atomic():
            conflicts = Person.objects.filter(source_url__in=seen_urls).count()
            if conflicts:
                raise CommandError(f"{conflicts} kaynak URL'si mevcut kayıtlarda var. Editör değişikliklerini korumak için aktarım durduruldu.")
            batch = ImportBatch.objects.create(filename=path.name, sha256=digest, row_count=len(records))
            for person in records:
                person.import_batch = batch
            Person.objects.bulk_create(records, batch_size=500)
        self.stdout.write(self.style.SUCCESS(f"{len(records)} şahıs aktarıldı; özgün tarih metinleri ve kaynak satırları korundu."))
