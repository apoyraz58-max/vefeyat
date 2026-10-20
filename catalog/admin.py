from django.contrib import admin, messages
from django.core.exceptions import ValidationError
from django.db import IntegrityError
from django.utils import timezone

from .models import Contribution, ImportBatch, NameVariant, Person


class VariantInline(admin.TabularInline):
    model = NameVariant
    extra = 0


@admin.register(Person)
class PersonAdmin(admin.ModelAdmin):
    list_display = ["short_name", "date_text", "source_url"]
    search_fields = ["short_name", "full_name", "original_name", "source_url"]
    readonly_fields = ["original_data", "source_sheet", "source_row", "import_batch", "updated_at"]
    exclude = ["short_norm", "full_norm", "original_norm"]
    inlines = [VariantInline]


@admin.register(Contribution)
class ContributionAdmin(admin.ModelAdmin):
    list_display = ["short_name", "kind", "status", "submitted_at", "reviewed_by"]
    list_filter = ["status", "kind"]
    search_fields = ["short_name", "note"]
    readonly_fields = ["status", "submitted_at", "reviewed_at", "reviewed_by"]
    autocomplete_fields = ["person"]
    actions = ["approve_selected", "reject_selected"]

    def get_readonly_fields(self, request, obj=None):
        fields = list(super().get_readonly_fields(request, obj))
        if obj and obj.status != Contribution.Status.PENDING:
            fields += ["kind", "person", "short_name", "date_text", "description", "source_url", "note"]
        return fields

    @admin.action(description="Seçili katkıları onayla ve yayımla", permissions=["change"])
    def approve_selected(self, request, queryset):
        count = 0
        for contribution in queryset:
            try:
                contribution.approve(request.user)
                count += 1
            except (ValidationError, IntegrityError, PermissionError) as error:
                self.message_user(request, f"{contribution.short_name}: {error}", messages.ERROR)
        self.message_user(request, f"{count} katkı onaylandı.", messages.SUCCESS)

    @admin.action(description="Seçili bekleyen katkıları reddet", permissions=["change"])
    def reject_selected(self, request, queryset):
        count = queryset.filter(status=Contribution.Status.PENDING).update(
            status=Contribution.Status.REJECTED, reviewed_by=request.user, reviewed_at=timezone.now()
        )
        self.message_user(request, f"{count} katkı reddedildi.")


@admin.register(ImportBatch)
class ImportBatchAdmin(admin.ModelAdmin):
    list_display = ["filename", "row_count", "imported_at"]
    readonly_fields = ["filename", "sha256", "row_count", "imported_at"]

    def has_add_permission(self, request):
        return False

    def has_delete_permission(self, request, obj=None):
        return False


admin.site.site_header = "Vefeyât · Editör"
admin.site.site_title = "Vefeyât"
admin.site.index_title = "Veri ve katkı yönetimi"
