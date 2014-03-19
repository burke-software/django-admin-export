from django.contrib.contenttypes.models import ContentType
from django.views.generic import TemplateView

from report_utils.mixins import GetFieldsMixin, DataExportMixin
from report_utils.model_introspection import get_relation_fields_from_model

class AdminExport(GetFieldsMixin, DataExportMixin, TemplateView):
    """ Get fields from a particular model """
    template_name = 'admin_export/export_to_xls.html'
    
    def get_context_data(self, **kwargs):
        context = super(AdminExport, self).get_context_data(**kwargs)
        field_name = self.request.GET.get('field', '')
        model_class = ContentType.objects.get(id=self.request.GET['ct']).model_class()
        if self.request.GET['ids'] == "IN_SESSION":
            queryset = model_class.objects.filter(pk__in=self.request.session['selected_ids'])
        else:
            queryset = model_class.objects.filter(pk__in=self.request.GET['ids'].split(','))
        path = self.request.GET.get('path', '')
        path_verbose = self.request.GET.get('path_verbose', '')
        context['model_name'] = model_class.__name__.lower()
        context['queryset'] = queryset
        context['model_ct'] = self.request.GET['ct']
        field_data = self.get_fields(model_class, field_name, path, path_verbose)
        context['related_fields'] = get_relation_fields_from_model(model_class)
        return dict(context.items() + field_data.items())

    def post(self, request, **kwargs):
        context = self.get_context_data(**kwargs)
        fields = []
        for field_name, value in request.POST.items():
            if value == "on":
                fields.append(field_name)
        data_list, message = self.report_to_list(
            context['queryset'],
            fields,
            self.request.user,
            )
        return self.list_to_xlsx_response(data_list, header=fields)
    
class AdminExportRelated(GetFieldsMixin, TemplateView):
    template_name = 'admin_export/fields.html'
    
    def post(self, request, **kwargs):
        context = self.get_context_data(**kwargs)
        model_class = ContentType.objects.get(id=self.request.POST['model_ct']).model_class()
        field_name = request.POST['field']
        path = request.POST['path']
        field_data = self.get_fields(model_class, field_name, path, '')
        context['related_fields'], model_ct, context['path'] = self.get_related_fields(model_class, field_name, path)
        context['model_ct'] = model_ct.id
        context['field_name'] = field_name
        context = dict(context.items() + field_data.items())
        return self.render_to_response(context)