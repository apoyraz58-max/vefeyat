import json

from django.conf import settings
from django.contrib.postgres.search import TrigramSimilarity
from django.core.exceptions import ValidationError
from django.core.paginator import Paginator
from django.db import connection
from django.db.models import Case, FloatField, Q, Value, When
from django.db.models.functions import Greatest
from django.http import JsonResponse
from django.shortcuts import get_object_or_404, render
from django.views.decorators.csrf import ensure_csrf_cookie
from django.views.decorators.http import require_GET, require_POST

from .models import Contribution, Person
from .search import normalize_name


def serialize_person(person, detail=False):
    result = {field: getattr(person, field) or None for field in [
        "short_name", "full_name", "original_name", "date_text", "description", "source_url",
    ]}
    result["id"] = str(person.pk)
    if detail:
        result["aliases"] = list(person.name_variants.values_list("name", flat=True))
    return result


@require_GET
@ensure_csrf_cookie
def index(request):
    return render(request, "catalog/index.html")


@require_GET
def health(request):
    with connection.cursor() as cursor:
        cursor.execute("SELECT 1")
        cursor.fetchone()
    return JsonResponse({"status": "ok", "database": "postgresql"})


@require_GET
def stats(request):
    return JsonResponse({"people_count": Person.objects.count(), "source_count": Person.objects.values("source_url").distinct().count()})


@require_GET
def people(request):
    query = request.GET.get("q", "").strip()
    if len(query) > 200:
        return JsonResponse({"error": "Arama en fazla 200 karakter olabilir."}, status=400)
    try:
        page_number = int(request.GET.get("page", 1))
        page_size = int(request.GET.get("page_size", 20))
        if page_number < 1 or not 1 <= page_size <= 50:
            raise ValueError
    except ValueError:
        return JsonResponse({"error": "Geçersiz sayfa; sayfa boyutu 1–50 arasında olmalı."}, status=400)
    queryset = Person.objects.all()
    normalized = normalize_name(query)
    approximate = False
    if query and len(normalized) < 2:
        return JsonResponse({"error": "En az iki harf veya rakam yazın."}, status=400)
    if normalized:
        direct = Q(short_norm__icontains=normalized) | Q(full_norm__icontains=normalized) | Q(original_norm__icontains=normalized) | Q(name_variants__normalized__icontains=normalized)
        exact = Q(short_norm=normalized) | Q(full_norm=normalized) | Q(original_norm=normalized) | Q(name_variants__normalized=normalized)
        if queryset.filter(exact).exists():
            match_filter = exact
        elif queryset.filter(direct).exists():
            match_filter = direct
        else:
            match_filter = Q(similarity__gte=0.25)
            approximate = True
        queryset = queryset.annotate(similarity=Greatest(
            TrigramSimilarity("short_norm", normalized),
            TrigramSimilarity("full_norm", normalized),
            TrigramSimilarity("original_norm", normalized),
        )).filter(
            match_filter
        ).annotate(rank=Case(
            When(short_norm=normalized, then=Value(3.0)),
            When(short_norm__startswith=normalized, then=Value(2.0)),
            When(short_norm__icontains=normalized, then=Value(1.0)),
            default=Value(0.0), output_field=FloatField(),
        )).order_by("-rank", "-similarity", "short_name", "id").distinct()
    paginator = Paginator(queryset, page_size)
    if page_number > max(paginator.num_pages, 1):
        return JsonResponse({"query": query, "approximate": approximate, "count": paginator.count, "page": page_number, "page_size": page_size, "results": [], "has_next": False})
    page = paginator.page(page_number)
    return JsonResponse({
        "query": query, "approximate": approximate, "count": paginator.count, "page": page.number, "page_size": page_size,
        "results": [serialize_person(person) for person in page.object_list], "has_next": page.has_next(),
    }, json_dumps_params={"ensure_ascii": False})


@require_GET
def person_detail(request, person_id):
    return JsonResponse(serialize_person(get_object_or_404(Person, pk=person_id), detail=True), json_dumps_params={"ensure_ascii": False})


@require_POST
def contribute(request):
    if request.content_type != "application/json":
        return JsonResponse({"error": "JSON içerik gereklidir."}, status=415)
    try:
        payload = json.loads(request.body)
        if not isinstance(payload, dict):
            raise ValueError
    except (ValueError, UnicodeDecodeError):
        return JsonResponse({"error": "Geçersiz JSON."}, status=400)
    lengths = {"kind": 20, "short_name": 500, "date_text": 500, "description": 5000, "source_url": 1000, "note": 5000}
    values = {}
    for field, limit in lengths.items():
        value = payload.get(field, "")
        if not isinstance(value, str) or len(value) > limit:
            return JsonResponse({"error": f"{field}: geçersiz değer veya uzunluk."}, status=400)
        values[field] = value.strip()
    person_id = payload.get("person_id")
    try:
        person = Person.objects.get(pk=person_id) if person_id else None
    except (ValidationError, Person.DoesNotExist, ValueError, TypeError):
        return JsonResponse({"error": "İlgili şahıs bulunamadı."}, status=400)
    contribution = Contribution(person=person, **values)
    try:
        contribution.full_clean()
    except ValidationError as error:
        return JsonResponse({"error": "Lütfen işaretli alanları kontrol edin.", "fields": error.message_dict}, status=400)
    contribution.save()
    return JsonResponse({"id": str(contribution.pk), "status": contribution.status, "message": "Öneriniz editör incelemesine alındı; henüz yayımlanmadı."}, status=201)


@require_GET
def word(request):
    response = render(request, "catalog/word.html")
    response["Content-Security-Policy"] = "frame-ancestors 'self' https://*.office.com https://*.officeapps.live.com https://*.microsoft365.com https://*.cloud.microsoft https://word.cloud.microsoft"
    return response


@require_GET
def manifest(request):
    response = render(request, "catalog/manifest.xml", {"base_url": settings.WORD_BASE_URL}, content_type="application/xml")
    response["Content-Disposition"] = 'attachment; filename="vefeyat-manifest.xml"'
    return response
