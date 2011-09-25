#   Copyright 2011 Burke Software and Consulting LLC
#   Author David M Burke <david@burkesoftware.com>
#   
# This file is part of django_admin_export.
#
# django_admin_export is free software: you can redistribute it and/or modify it 
# under theterms of the GNU General Public License as published by the Free 
# Software Foundation, either version 3 of the License, or (at your option) any 
# later version.
#
# django_admin_export is distributed in the hope that it will be useful, but 
# WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or 
# FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
# details.
#
# You should have received a copy of the GNU General Public License along with
# django_admin_export. If not, see http://www.gnu.org/licenses/.

from django.contrib import admin
from django.contrib.contenttypes.models import ContentType
from django.http import HttpResponseRedirect

def export_simple_selected_objects(modeladmin, request, queryset):
    selected_int = queryset.values_list('id', flat=True)
    selected = []
    for s in selected_int:
        selected.append(str(s))
    ct = ContentType.objects.get_for_model(queryset.model)
    return HttpResponseRedirect("/admin_export/export_to_xls/?ct=%s&ids=%s" % (ct.pk, ",".join(selected)))
export_simple_selected_objects.short_description = "Export selected items to XLS"
admin.site.add_action(export_simple_selected_objects)