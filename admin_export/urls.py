from django.conf.urls.defaults import *
from views import *

urlpatterns = patterns('',
    (r'^export_to_xls/$', admin_export_xls),
    (r'^export_to_xls_related/$', get_fields_for_model),
)
