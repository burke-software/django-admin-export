# -- encoding: UTF-8 --
import random
from admin_export.views import AdminExport
from django.contrib import admin
from django.contrib.contenttypes.models import ContentType
from django.core.urlresolvers import reverse
from django.utils.http import urlencode
import pytest
from .models import ModelUnderTest, ModelWithRelated


class ModelAdminTest(admin.ModelAdmin):
    def get_queryset(self, request):
        return super(ModelAdminTest, self).get_queryset(request).filter(value__lt=request.magic)


def queryset_valid(request, queryset):
    return all(x.value < request.magic for x in queryset)


@pytest.mark.django_db
def test_queryset_from_admin(rf, admin_user):
    for x in range(100):
        ModelUnderTest.objects.get_or_create(value=x)
    assert ModelUnderTest.objects.count() >= 100

    request = rf.get("/")
    request.user = admin_user
    request.magic = random.randint(10, 90)
    request.GET = {
        "ct": ContentType.objects.get_for_model(ModelUnderTest).pk,
        "ids": ",".join(str(id) for id in ModelUnderTest.objects.all().values_list("pk", flat=True))
    }

    old_registry = admin.site._registry
    admin.site._registry = {}
    admin.site.register(ModelUnderTest, ModelAdminTest)
    assert queryset_valid(request, admin.site._registry[ModelUnderTest].get_queryset(request))
    assert not queryset_valid(request, ModelUnderTest.objects.all())

    admin_export_view = AdminExport()
    admin_export_view.request = request
    admin_export_view.args = ()
    admin_export_view.kwargs = {}
    assert admin_export_view.get_model_class() == ModelUnderTest
    assert queryset_valid(request, admin_export_view.get_queryset(ModelUnderTest))
    admin.site._registry = old_registry


@pytest.mark.django_db
@pytest.mark.parametrize('output_name', ['html', 'csv'])
def test_AdminExport_list_to_method_response_should_return_200(admin_user, output_name):
    for x in range(3):
        ModelUnderTest.objects.get_or_create(value=x)

    admin = AdminExport()
    data = admin.report_to_list(ModelUnderTest.objects.all(), ['value'], admin_user)

    method = getattr(admin, 'list_to_{}_response'.format(output_name))
    res = method(data)
    assert res.status_code == 200


@pytest.mark.django_db
@pytest.mark.parametrize('output_format', ['html', 'csv', 'xls'])
def test_AdminExport_post_should_return_200(admin_client, output_format):
    for x in range(3):
        ModelUnderTest.objects.get_or_create(value=x)

    params = {
        'ct': ContentType.objects.get_for_model(ModelUnderTest).pk,
        'ids': ','.join(repr(pk) for pk in ModelUnderTest.objects.values_list('pk', flat=True))
    }
    data = {
        "value": "on",
        "__format": output_format,
    }
    url = "{}?{}".format(reverse('admin_export:export'), urlencode(params))
    response = admin_client.post(url, data=data)
    assert response.status_code == 200


@pytest.mark.django_db
def test_AdminExport_get_should_return_200(admin_client):
    for x in range(3):
        ModelUnderTest.objects.get_or_create(value=x)

    params = {
        'ct': ContentType.objects.get_for_model(ModelUnderTest).pk,
        'ids': ','.join(repr(pk) for pk in ModelUnderTest.objects.values_list('pk', flat=True))
    }
    url = "{}?{}".format(reverse('admin_export:export'), urlencode(params))
    response = admin_client.get(url)
    assert response.status_code == 200


@pytest.mark.django_db
def test_AdminExport_with_related_get_should_return_200(admin_client):
    mut, created = ModelUnderTest.objects.get_or_create(value=1)
    for x in range(3):
        ModelWithRelated.objects.get_or_create(value=x, mut=mut)

    params = {
        'related': True,
        'model_ct': ContentType.objects.get_for_model(ModelWithRelated).pk,
        'field': 'mut',
        'path': 'mut.value',
    }
    url = "{}?{}".format(reverse('admin_export:export'), urlencode(params))
    response = admin_client.get(url)
    assert response.status_code == 200


@pytest.mark.django_db
def test_AdminExport_with_unregistered_model_should_raise_ValueError(admin_client):
    mut, created = ModelUnderTest.objects.get_or_create(value=1)
    for x in range(3):
        ModelWithRelated.objects.get_or_create(value=x, mut=mut)

    params = {
        'ct': ContentType.objects.get_for_model(ModelWithRelated).pk,
        'ids': ','.join(repr(pk) for pk in ModelWithRelated.objects.values_list('pk', flat=True))
    }
    url = "{}?{}".format(reverse('admin_export:export'), urlencode(params))

    with pytest.raises(ValueError):
        admin_client.get(url)
