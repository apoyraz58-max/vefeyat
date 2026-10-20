from pathlib import Path
from urllib.parse import urlsplit

from django.core.management.base import BaseCommand, CommandError
from django.template.loader import render_to_string


class Command(BaseCommand):
    help = "HTTPS yayın adresi için Windows Word eklenti manifestini üretir."

    def add_arguments(self, parser):
        parser.add_argument("--base-url", required=True)
        parser.add_argument("--output", type=Path, required=True)

    def handle(self, *args, **options):
        base = options["base_url"].rstrip("/")
        parsed = urlsplit(base)
        if (parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password
                or parsed.path or parsed.query or parsed.fragment):
            raise CommandError("Kullanıcı bilgisi, yol veya sorgu içermeyen bir HTTPS kök adresi verin.")
        text = render_to_string("catalog/manifest.xml", {"base_url": base})
        output = options["output"]
        if output.exists():
            raise CommandError("Hedef dosya zaten var; korumak için üzerine yazılmadı.")
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open("x", encoding="utf-8") as stream:
            stream.write(text)
        self.stdout.write(self.style.SUCCESS("Word manifesti oluşturuldu."))
