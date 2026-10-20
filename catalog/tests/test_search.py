from django.test import TestCase
from django.urls import reverse

from catalog.models import NameVariant, Person


class PeopleSearchTests(TestCase):
    @classmethod
    def setUpTestData(cls):
        cls.sina = Person.objects.create(
            short_name="İbn Sînâ", full_name="Ebû Alî el-Hüseyn b. Abdullah İbn Sînâ",
            original_name="ابن سينا", date_text="(ö. 428/1037)",
            description="Filozof ve hekim.", source_url="https://example.org/ibn-sina",
        )
        NameVariant.objects.create(person=cls.sina, name="Avicenna")
        NameVariant.objects.create(person=cls.sina, name="Avicenna Latinus")
        cls.rusd_elder = Person.objects.create(
            short_name="İbn Rüşd", date_text="(ö. 520/1126)",
            description="Fakih.", source_url="https://example.org/rusd-elder",
        )
        cls.rusd_younger = Person.objects.create(
            short_name="İbn Rüşd", date_text="(ö. 595/1198)",
            description="Filozof.", source_url="https://example.org/rusd-younger",
        )

    def results(self, query, **params):
        response = self.client.get(reverse("people"), {"q": query, **params})
        self.assertEqual(response.status_code, 200)
        return response.json()

    def test_ibni_and_turkish_diacritics_find_preserved_display_name(self):
        for query in ("ibni sina", "İBNİ SİNA", "ibn sînâ"):
            with self.subTest(query=query):
                result = self.results(query)
                self.assertEqual(result["results"][0]["id"], str(self.sina.pk))
                self.assertEqual(result["results"][0]["short_name"], "İbn Sînâ")
                self.assertEqual(result["results"][0]["date_text"], "(ö. 428/1037)")

    def test_misspelling_returns_a_useful_fuzzy_candidate(self):
        result = self.results("ibn sinaa")
        self.assertTrue(result["approximate"])
        self.assertIn(str(self.sina.pk), [person["id"] for person in result["results"]])

    def test_exact_name_excludes_longer_names_and_similar_spelling(self):
        Person.objects.create(
            short_name="İbn Sînâ el-Belhî", description="Benzer isimli ayrı bir şahıs.",
            source_url="https://example.org/sina-longer",
        )
        Person.objects.create(
            short_name="İbn Sinan", description="Yazımı yakın, farklı bir şahıs.",
            source_url="https://example.org/sinan",
        )
        result = self.results("ibni sina")
        self.assertFalse(result["approximate"])
        self.assertEqual(result["count"], 1)
        self.assertEqual([person["id"] for person in result["results"]], [str(self.sina.pk)])

    def test_direct_partial_match_excludes_fuzzy_only_candidates(self):
        Person.objects.create(
            short_name="İbn Sine", description="Yazımı yakın, farklı bir şahıs.",
            source_url="https://example.org/sine",
        )
        result = self.results("sina")
        self.assertFalse(result["approximate"])
        self.assertEqual(result["count"], 1)
        self.assertEqual([person["id"] for person in result["results"]], [str(self.sina.pk)])

    def test_original_arabic_name_is_searchable(self):
        result = self.results("ابن سينا")
        self.assertEqual(result["results"][0]["id"], str(self.sina.pk))
        self.assertEqual(result["results"][0]["original_name"], "ابن سينا")

    def test_alias_search_does_not_duplicate_a_person(self):
        result = self.results("Avicenna")
        self.assertEqual(result["count"], 1)
        self.assertEqual([person["id"] for person in result["results"]], [str(self.sina.pk)])

    def test_people_with_same_name_remain_distinguishable(self):
        result = self.results("ibn rusd")
        matches = [person for person in result["results"] if person["short_name"] == "İbn Rüşd"]
        self.assertEqual({person["id"] for person in matches}, {str(self.rusd_elder.pk), str(self.rusd_younger.pk)})
        self.assertEqual({person["date_text"] for person in matches}, {"(ö. 520/1126)", "(ö. 595/1198)"})
        self.assertEqual(len({person["source_url"] for person in matches}), 2)

    def test_invalid_search_parameters_return_client_errors(self):
        for params in (
            {"q": "x"}, {"q": "!!!"}, {"q": "a" * 201},
            {"page": "wrong"}, {"page": "0"}, {"page": "-1"},
            {"page_size": "x"}, {"page_size": "0"}, {"page_size": "51"},
        ):
            with self.subTest(params=params):
                response = self.client.get(reverse("people"), params)
                self.assertEqual(response.status_code, 400)

    def test_pagination_is_stable_and_out_of_range_is_empty(self):
        first = self.results("", page_size=1)
        second = self.results("", page_size=1, page=2)
        self.assertEqual(first["count"], 3)
        self.assertTrue(first["has_next"])
        self.assertNotEqual(first["results"][0]["id"], second["results"][0]["id"])
        beyond = self.results("", page_size=1, page=99)
        self.assertEqual(beyond["results"], [])
        self.assertEqual(beyond["count"], 3)
        self.assertFalse(beyond["has_next"])

    def test_detail_includes_aliases_and_unknown_uuid_is_404(self):
        response = self.client.get(reverse("person_detail", args=[self.sina.pk]))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(set(response.json()["aliases"]), {"Avicenna", "Avicenna Latinus"})
        missing = self.client.get(reverse("person_detail", args=["00000000-0000-0000-0000-000000000000"]))
        self.assertEqual(missing.status_code, 404)
