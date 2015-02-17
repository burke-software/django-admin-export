from django.conf.urls import url, patterns
from django.contrib.admin.views.decorators import staff_member_required
from .views import AdminExport

view = staff_member_required(AdminExport.as_view())
urlpatterns = patterns('',
   url(r'^export/$', view, name="export"),
   (r'^export_to_xls/$', view),  # compatibility for users who upgrade without touching URLs
)
