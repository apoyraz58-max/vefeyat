import uuid

from django.conf import settings
from django.contrib.postgres.indexes import GinIndex
from django.core.exceptions import ValidationError
from django.db import models, transaction
from django.utils import timezone

from .search import normalize_name


class ImportBatch(models.Model):
    filename = models.CharField("Dosya adı", max_length=255)
    sha256 = models.CharField("SHA-256", max_length=64, unique=True)
    row_count = models.PositiveIntegerField(default=0)
    imported_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.filename} · {self.row_count} kayıt"


class Person(models.Model):
    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    short_name = models.CharField("Kısa ad", max_length=500)
    full_name = models.TextField("Tam ad", blank=True)
    original_name = models.TextField("Orijinal ad", blank=True)
    date_text = models.CharField("Kaynaktaki tarih ifadesi", max_length=500, blank=True)
    description = models.TextField("Tanıtım", blank=True)
    source_url = models.URLField("Birincil kaynak", max_length=1000, unique=True)
    short_norm = models.TextField(editable=False)
    full_norm = models.TextField(editable=False, blank=True)
    original_norm = models.TextField(editable=False, blank=True)
    original_data = models.JSONField("Özgün aktarım verisi", default=dict, editable=False)
    source_sheet = models.CharField(max_length=200, blank=True, editable=False)
    source_row = models.PositiveIntegerField(null=True, blank=True, editable=False)
    import_batch = models.ForeignKey(ImportBatch, on_delete=models.SET_NULL, null=True, blank=True, editable=False)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["short_name", "id"]
        verbose_name = "Şahıs"
        verbose_name_plural = "Şahıslar"
        indexes = [GinIndex(fields=["short_norm"], opclasses=["gin_trgm_ops"], name="person_short_trgm")]

    def save(self, *args, **kwargs):
        self.short_norm = normalize_name(self.short_name)
        self.full_norm = normalize_name(self.full_name)
        self.original_norm = normalize_name(self.original_name)
        if kwargs.get("update_fields") is not None:
            kwargs["update_fields"] = set(kwargs["update_fields"]) | {"short_norm", "full_norm", "original_norm"}
        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.short_name} {self.date_text}".strip()


class NameVariant(models.Model):
    person = models.ForeignKey(Person, related_name="name_variants", on_delete=models.CASCADE)
    name = models.CharField("İsim varyantı", max_length=500)
    normalized = models.TextField(editable=False)

    class Meta:
        verbose_name = "İsim varyantı"
        verbose_name_plural = "İsim varyantları"
        constraints = [models.UniqueConstraint(fields=["person", "name"], name="unique_person_variant")]

    def save(self, *args, **kwargs):
        self.normalized = normalize_name(self.name)
        super().save(*args, **kwargs)

    def __str__(self):
        return self.name


class Contribution(models.Model):
    class Kind(models.TextChoices):
        NEW = "new_person", "Yeni şahıs önerisi"
        CORRECTION = "correction", "Düzeltme önerisi"

    class Status(models.TextChoices):
        PENDING = "pending", "İnceleme bekliyor"
        APPROVED = "approved", "Onaylandı"
        REJECTED = "rejected", "Reddedildi"

    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    kind = models.CharField("Tür", max_length=20, choices=Kind.choices)
    person = models.ForeignKey(Person, on_delete=models.PROTECT, null=True, blank=True, verbose_name="İlgili şahıs")
    short_name = models.CharField("Önerilen kısa ad", max_length=500)
    date_text = models.CharField("Önerilen tarih ifadesi", max_length=500, blank=True)
    description = models.TextField("Önerilen tanıtım", max_length=5000, blank=True)
    source_url = models.URLField("Kaynak", max_length=1000)
    note = models.TextField("Gerekçe / açıklama", max_length=5000)
    status = models.CharField("Durum", max_length=20, choices=Status.choices, default=Status.PENDING)
    submitted_at = models.DateTimeField("Gönderim", auto_now_add=True)
    reviewed_at = models.DateTimeField("İnceleme tarihi", null=True, blank=True, editable=False)
    reviewed_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, editable=False)
    review_note = models.TextField("Editör notu", blank=True)

    class Meta:
        ordering = ["-submitted_at"]
        verbose_name = "Katkı"
        verbose_name_plural = "Katkılar"

    def clean(self):
        if self.kind == self.Kind.CORRECTION and not self.person_id:
            raise ValidationError({"person": "Düzeltme için bir şahıs seçilmelidir."})
        if self.kind == self.Kind.NEW and self.person_id and self.status == self.Status.PENDING:
            raise ValidationError({"person": "Yeni şahıs önerisi mevcut kayda bağlanamaz."})
        if not self.note.strip():
            raise ValidationError({"note": "Gerekçe yazılmalıdır."})
        if not self.source_url.startswith("https://"):
            raise ValidationError({"source_url": "HTTPS kaynak bağlantısı gereklidir."})

    def approve(self, reviewer):
        """Publish only through an explicit, permission-checked review action."""
        if not reviewer.is_active or not reviewer.has_perm("catalog.change_contribution") or not reviewer.has_perm("catalog.change_person"):
            raise PermissionError("Katkı inceleme ve şahıs düzenleme yetkisi gerekir.")
        with transaction.atomic():
            locked = type(self).objects.select_for_update().get(pk=self.pk)
            if locked.status != self.Status.PENDING:
                raise ValidationError("Bu katkı daha önce incelenmiş.")
            if locked.kind == self.Kind.NEW and not reviewer.has_perm("catalog.add_person"):
                raise PermissionError("Yeni şahıs ekleme yetkisi gerekir.")
            locked.full_clean()
            if locked.kind == self.Kind.NEW:
                target = Person()
            else:
                target = Person.objects.select_for_update().get(pk=locked.person_id)
            target.short_name = locked.short_name
            # Blank optional fields mean 'no proposed change' for corrections.
            if locked.date_text or locked.kind == self.Kind.NEW:
                target.date_text = locked.date_text
            if locked.description or locked.kind == self.Kind.NEW:
                target.description = locked.description
            target.source_url = locked.source_url
            target.full_clean(exclude=["short_norm", "full_norm", "original_norm", "original_data"])
            target.save()
            locked.person = target
            locked.status = self.Status.APPROVED
            locked.reviewed_by = reviewer
            locked.reviewed_at = timezone.now()
            locked.save(update_fields=["person", "status", "reviewed_by", "reviewed_at"])
        self.refresh_from_db()
        return target

    def __str__(self):
        return f"{self.get_kind_display()}: {self.short_name}"
