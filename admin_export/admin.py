from django.contrib import admin
from django.contrib.contenttypes.models import ContentType
from django.http import HttpResponseRedirect


def export_simple_selected_objects(modeladmin, request, queryset):
    selected_int = queryset.values_list('id', flat=True)
    selected = []
    for s in selected_int:
        selected.append(str(s))
    ct = ContentType.objects.get_for_model(queryset.model)
    if len(selected) > 10000:
        request.session['selected_ids'] = selected
        return HttpResponseRedirect("/admin_export/export_to_xls/?ct=%s&ids=IN_SESSION" % (ct.pk,))
    else:
        return HttpResponseRedirect("/admin_export/export_to_xls/?ct=%s&ids=%s" % (ct.pk, ",".join(selected)))
export_simple_selected_objects.short_description = "Export selected items to XLS"
admin.site.add_action(export_simple_selected_objects)
