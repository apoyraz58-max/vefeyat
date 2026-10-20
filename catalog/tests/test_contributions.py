import json

from django.contrib.auth import get_user_model
from django.contrib.auth.models import Permission
from django.core.exceptions import ValidationError
from django.test import Client, TestCase
from django.urls import reverse

from catalog.models import Contribution, Person


def give_review_permissions(user, *codenames):
    user.user_permissions.add(*Permission.objects.filter(content_type__app_label="catalog", codename__in=codenames))
    return user


class PublicContributionTests(TestCase):
    @classmethod
    def setUpTestData(cls):
        cls.person = Person.objects.create(
            short_name="İbn Sînâ", date_text="(ö. 428/1037)",
            description="Kaynakta yer alan tanıtım.", source_url="https://example.org/ibn-sina",
            original_data={"Kısa Ad": "İbn Sînâ", "Vefat Tarihi": "(ö. 428/1037)"},
        )

    def setUp(self):
        self.browser = Client(enforce_csrf_checks=True)
        self.payload = {
            "kind": "correction", "person_id": str(self.person.pk),
            "short_name": "İbn Sînâ", "date_text": "(ö. 428/1037)",
            "description": "Daha ayrıntılı bir tanıtım önerisi.",
            "source_url": "https://example.org/ibn-sina",
            "note": "Tanıtımın kaynaktaki açıklamaya göre genişletilmesi öneriliyor.",
        }

    def csrf_post(self, payload):
        homepage = self.browser.get(reverse("index"))
        self.assertEqual(homepage.status_code, 200)
        self.assertIn("csrftoken", self.browser.cookies)
        return self.browser.post(
            reverse("contribute"), data=json.dumps(payload), content_type="application/json",
            HTTP_X_CSRFTOKEN=self.browser.cookies["csrftoken"].value,
        )

    def test_submission_without_csrf_token_is_rejected(self):
        response = self.browser.post(reverse("contribute"), data=json.dumps(self.payload), content_type="application/json")
        self.assertEqual(response.status_code, 403)
        self.assertFalse(Contribution.objects.exists())

    def test_public_correction_stays_pending_and_does_not_modify_person(self):
        # A visitor cannot use extra JSON fields to publish their own suggestion.
        self.payload.update(status="approved", reviewed_by=1)
        response = self.csrf_post(self.payload)
        self.assertEqual(response.status_code, 201)
        self.assertEqual(response.json()["status"], "pending")
        contribution = Contribution.objects.get()
        self.assertEqual(contribution.status, Contribution.Status.PENDING)
        self.assertIsNone(contribution.reviewed_by)
        self.assertIsNone(contribution.reviewed_at)
        self.person.refresh_from_db()
        self.assertEqual(self.person.description, "Kaynakta yer alan tanıtım.")
        self.assertEqual(Person.objects.count(), 1)

    def test_new_person_submission_is_not_published_before_review(self):
        self.payload.update(kind="new_person", short_name="Yeni araştırılan şahıs", source_url="https://example.org/new")
        self.payload.pop("person_id")
        response = self.csrf_post(self.payload)
        self.assertEqual(response.status_code, 201)
        contribution = Contribution.objects.get()
        self.assertIsNone(contribution.person_id)
        self.assertEqual(contribution.status, Contribution.Status.PENDING)
        self.assertEqual(Person.objects.count(), 1)

    def test_unsafe_source_urls_are_rejected(self):
        for url in ("http://example.org/source", "javascript:alert(1)", "file:///etc/passwd", "not-a-url"):
            with self.subTest(url=url):
                response = self.csrf_post({**self.payload, "source_url": url})
                self.assertEqual(response.status_code, 400)
                self.assertIn("source_url", response.json()["fields"])
        self.assertFalse(Contribution.objects.exists())

    def test_correction_needs_an_existing_person(self):
        for person_id in (None, "bad-uuid", "00000000-0000-0000-0000-000000000000"):
            with self.subTest(person_id=person_id):
                response = self.csrf_post({**self.payload, "person_id": person_id})
                self.assertEqual(response.status_code, 400)
        self.assertFalse(Contribution.objects.exists())

    def test_new_person_cannot_target_an_existing_record(self):
        response = self.csrf_post({**self.payload, "kind": "new_person"})
        self.assertEqual(response.status_code, 400)
        self.assertFalse(Contribution.objects.exists())

    def test_malformed_payloads_are_client_errors_and_create_nothing(self):
        # Pass through real CSRF middleware for JSON failures, too.
        self.browser.get(reverse("index"))
        token = self.browser.cookies["csrftoken"].value
        for body in ('{"broken":', "[]", "null"):
            with self.subTest(body=body):
                response = self.browser.post(
                    reverse("contribute"), data=body, content_type="application/json", HTTP_X_CSRFTOKEN=token,
                )
                self.assertEqual(response.status_code, 400)
        for change in ({"short_name": ["a"]}, {"note": " "}, {"description": "x" * 5001}):
            with self.subTest(fields=list(change)):
                self.assertEqual(self.csrf_post({**self.payload, **change}).status_code, 400)
        self.assertFalse(Contribution.objects.exists())


class ContributionReviewTests(TestCase):
    @classmethod
    def setUpTestData(cls):
        cls.reviewer = give_review_permissions(
            get_user_model().objects.create_user(username="editor", is_staff=True),
            "change_contribution", "change_person", "add_person",
        )
        cls.original = {"Kısa Ad": "İbn Sînâ", "Tam Ad": "---", "Vefat Tarihi": "(ö. 428/1037)"}
        cls.person = Person.objects.create(
            short_name="İbn Sînâ", date_text="(ö. 428/1037)",
            description="Özgün tanıtım.", source_url="https://example.org/ibn-sina",
            original_data=cls.original, source_sheet="alimler_251025", source_row=14,
        )

    def correction(self, **overrides):
        values = {
            "kind": Contribution.Kind.CORRECTION, "person": self.person,
            "short_name": "İbn Sînâ", "date_text": "(ö. 428/1037)",
            "description": "Editörün doğruladığı yeni tanıtım.",
            "source_url": self.person.source_url, "note": "Kaynakla karşılaştırıldı.",
        }
        values.update(overrides)
        return Contribution.objects.create(**values)

    def new_person(self):
        return Contribution.objects.create(
            kind=Contribution.Kind.NEW, short_name="Yeni şahıs", date_text="(tarih bilinmiyor)",
            source_url="https://example.org/new-person", note="Kaynak kontrolü yapıldı.",
        )

    def test_authorized_approval_publishes_correction_and_preserves_raw_provenance(self):
        contribution = self.correction()
        target = contribution.approve(self.reviewer)
        self.assertEqual(target.pk, self.person.pk)
        self.person.refresh_from_db()
        self.assertEqual(self.person.description, "Editörün doğruladığı yeni tanıtım.")
        self.assertEqual(self.person.original_data, self.original)
        self.assertEqual(self.person.source_sheet, "alimler_251025")
        self.assertEqual(self.person.source_row, 14)
        contribution.refresh_from_db()
        self.assertEqual(contribution.status, Contribution.Status.APPROVED)
        self.assertEqual(contribution.reviewed_by, self.reviewer)
        self.assertIsNotNone(contribution.reviewed_at)
        with self.assertRaises(ValidationError):
            contribution.approve(self.reviewer)
        self.assertEqual(Person.objects.count(), 1)

    def test_blank_optional_correction_fields_preserve_existing_information(self):
        self.correction(date_text="", description="").approve(self.reviewer)
        self.person.refresh_from_db()
        self.assertEqual(self.person.date_text, "(ö. 428/1037)")
        self.assertEqual(self.person.description, "Özgün tanıtım.")

    def test_new_person_is_published_once(self):
        contribution = self.new_person()
        target = contribution.approve(self.reviewer)
        self.assertEqual(Person.objects.count(), 2)
        self.assertEqual(target.short_name, "Yeni şahıs")
        contribution.refresh_from_db()
        self.assertEqual(contribution.person_id, target.pk)
        self.assertEqual(contribution.status, Contribution.Status.APPROVED)
        with self.assertRaises(ValidationError):
            contribution.approve(self.reviewer)
        self.assertEqual(Person.objects.count(), 2)

    def test_review_requires_person_and_contribution_permissions(self):
        contribution = self.correction()
        permission_sets = ((), ("change_contribution",), ("change_person",))
        for index, permissions in enumerate(permission_sets):
            with self.subTest(permissions=permissions):
                user = give_review_permissions(
                    get_user_model().objects.create_user(username=f"limited-{index}", is_staff=True), *permissions,
                )
                with self.assertRaises(PermissionError):
                    contribution.approve(user)
                contribution.refresh_from_db()
                self.assertEqual(contribution.status, Contribution.Status.PENDING)
        self.person.refresh_from_db()
        self.assertEqual(self.person.description, "Özgün tanıtım.")

    def test_new_person_requires_add_permission(self):
        user = give_review_permissions(
            get_user_model().objects.create_user(username="correction-editor", is_staff=True),
            "change_contribution", "change_person",
        )
        contribution = self.new_person()
        with self.assertRaises(PermissionError):
            contribution.approve(user)
        contribution.refresh_from_db()
        self.assertEqual(contribution.status, Contribution.Status.PENDING)
        self.assertEqual(Person.objects.count(), 1)

    def test_add_permission_uses_current_contribution_kind_not_a_stale_instance(self):
        user = give_review_permissions(
            get_user_model().objects.create_user(username="stale-correction-editor", is_staff=True),
            "change_contribution", "change_person",
        )
        stale_contribution = self.correction()
        # Another editor changes the pending suggestion after the first editor
        # loaded it. Publication must authorize the kind actually being saved.
        Contribution.objects.filter(pk=stale_contribution.pk).update(
            kind=Contribution.Kind.NEW, person=None, source_url="https://example.org/changed-kind",
        )
        with self.assertRaises(PermissionError):
            stale_contribution.approve(user)
        stale_contribution.refresh_from_db()
        self.assertEqual(stale_contribution.status, Contribution.Status.PENDING)
        self.assertEqual(Person.objects.count(), 1)

    def test_get_cannot_trigger_admin_approval(self):
        contribution = self.correction()
        self.client.force_login(self.reviewer)
        response = self.client.get(reverse("admin:catalog_contribution_changelist"), {
            "action": "approve_selected", "_selected_action": str(contribution.pk),
        })
        self.assertIn(response.status_code, (200, 302))
        contribution.refresh_from_db()
        self.assertEqual(contribution.status, Contribution.Status.PENDING)
        self.person.refresh_from_db()
        self.assertEqual(self.person.description, "Özgün tanıtım.")

    def test_anonymous_admin_post_cannot_approve(self):
        contribution = self.correction()
        response = self.client.post(reverse("admin:catalog_contribution_changelist"), {
            "action": "approve_selected", "_selected_action": str(contribution.pk), "index": "0",
        })
        self.assertEqual(response.status_code, 302)
        self.assertIn(reverse("admin:login"), response.url)
        contribution.refresh_from_db()
        self.assertEqual(contribution.status, Contribution.Status.PENDING)

    def test_staff_with_contribution_permission_cannot_bypass_person_permissions_via_admin(self):
        user = give_review_permissions(
            get_user_model().objects.create_user(username="contribution-only", is_staff=True), "change_contribution",
        )
        contribution = self.correction()
        self.client.force_login(user)
        response = self.client.post(reverse("admin:catalog_contribution_changelist"), {
            "action": "approve_selected", "_selected_action": str(contribution.pk), "index": "0",
        })
        self.assertEqual(response.status_code, 302)
        contribution.refresh_from_db()
        self.assertEqual(contribution.status, Contribution.Status.PENDING)
        self.person.refresh_from_db()
        self.assertEqual(self.person.description, "Özgün tanıtım.")

    def test_admin_post_approves_with_full_permissions(self):
        contribution = self.new_person()
        self.client.force_login(self.reviewer)
        response = self.client.post(reverse("admin:catalog_contribution_changelist"), {
            "action": "approve_selected", "_selected_action": str(contribution.pk), "index": "0",
        })
        self.assertEqual(response.status_code, 302)
        contribution.refresh_from_db()
        self.assertEqual(contribution.status, Contribution.Status.APPROVED)
        self.assertEqual(Person.objects.count(), 2)
