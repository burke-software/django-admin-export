# -- encoding: UTF-8 --
import random
from admin_export.views import AdminExport
from django.contrib import admin
from django.contrib.contenttypes.models import ContentType
import pytest
from tests.models import TestModel


class TestModelAdmin(admin.ModelAdmin):
    def get_queryset(self, request):
        return super(TestModelAdmin, self).get_queryset(request).filter(value__lt=request.magic)


def queryset_valid(request, queryset):
    return all(x.value < request.magic for x in queryset)


@pytest.mark.django_db
def test_queryset_from_admin(rf, admin_user):
    for x in range(100):
        TestModel.objects.get_or_create(value=x)
    assert TestModel.objects.count() >= 100

    request = rf.get("/")
    request.user = admin_user
    request.magic = random.randint(10, 90)
    request.GET = {
        "ct": ContentType.objects.get_for_model(TestModel).pk,
        "ids": ",".join(str(id) for id in TestModel.objects.all().values_list("pk", flat=True))
    }

    old_registry = admin.site._registry
    admin.site._registry = {}
    admin.site.register(TestModel, TestModelAdmin)
    assert queryset_valid(request, admin.site._registry[TestModel].get_queryset(request))
    assert not queryset_valid(request, TestModel.objects.all())

    admin_export_view = AdminExport()
    admin_export_view.request = request
    admin_export_view.args = ()
    admin_export_view.kwargs = {}
    assert admin_export_view.get_model_class() == TestModel
    assert queryset_valid(request, admin_export_view.get_queryset(TestModel))
    admin.site._registry = old_registry
