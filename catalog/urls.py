from django.urls import path
from . import views

urlpatterns = [
    path("", views.index, name="index"),
    path("api/health/", views.health, name="health"),
    path("api/stats/", views.stats, name="stats"),
    path("api/people/", views.people, name="people"),
    path("api/people/<uuid:person_id>/", views.person_detail, name="person_detail"),
    path("api/contributions/", views.contribute, name="contribute"),
    path("word/", views.word, name="word"),
    path("word/manifest.xml", views.manifest, name="manifest"),
]
