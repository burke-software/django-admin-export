from itertools import chain
from django.contrib import admin
from django.contrib.contenttypes.models import ContentType
from django.http.response import HttpResponse
from django.template import loader, Context
from django.views.generic import TemplateView
import csv
from report_utils.mixins import GetFieldsMixin, DataExportMixin
from report_utils.model_introspection import get_relation_fields_from_model

HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{{ title }}</title>
</head>
<body>
    <h1>{{ title }}</h1>
    <table border=1>
    {% if header %}<thead><tr>{% for h in header %}<th>{{ h }}</th>{% endfor %}</tr></thead>{% endif %}
    <tbody>
        {% for datum in data %}
        <tr>{% for cell in datum %}<td>{{ cell|linebreaksbr }}</td>{% endfor %}</tr>
        {% endfor %}
    </tbody>
    </table>
</body>
</html>
"""


class ExtDataExportMixin(DataExportMixin):

    def list_to_html_response(self, data, title='', header=None):
        html = loader.get_template_from_string(HTML_TEMPLATE).render(Context(locals()))
        return HttpResponse(html)

    def list_to_csv_response(self, data, title='', header=None):
        resp = HttpResponse(content_type="text/csv; charset=UTF-8")
        cw = csv.writer(resp)
        for row in chain([header] if header else [], data):
            cw.writerow([unicode(s).encode(resp._charset) for s in row])
        return resp


class AdminExport(GetFieldsMixin, ExtDataExportMixin, TemplateView):

    """ Get fields from a particular model """
    template_name = 'admin_export/export.html'

    def get_context_data(self, **kwargs):
        context = super(AdminExport, self).get_context_data(**kwargs)
        field_name = self.request.GET.get('field', '')
        model_class = ContentType.objects.get(id=self.request.GET['ct']).model_class()

        if self.request.GET.get("session_key"):
            ids = self.request.session[self.request.GET["session_key"]]
        else:
            ids = self.request.GET['ids'].split(',')

        queryset = model_class.objects.filter(pk__in=ids)

        path = self.request.GET.get('path', '')
        path_verbose = self.request.GET.get('path_verbose', '')
        context['opts'] = model_class._meta
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
        format = request.POST.get("__format")
        if format == "html":
            return self.list_to_html_response(data_list, header=fields)
        elif format == "csv":
            return self.list_to_csv_response(data_list, header=fields)
        else:
            return self.list_to_xlsx_response(data_list, header=fields)

    def get(self, request, *args, **kwargs):
        if request.REQUEST.get("related"):  # Dispatch to the other view
            return AdminExportRelated.as_view()(request=self.request)
        return super(AdminExport, self).get(request, *args, **kwargs)


class AdminExportRelated(GetFieldsMixin, TemplateView):
    template_name = 'admin_export/fields.html'

    def get(self, request, **kwargs):
        context = self.get_context_data(**kwargs)
        model_class = ContentType.objects.get(id=self.request.GET['model_ct']).model_class()
        field_name = request.GET['field']
        path = request.GET['path']
        field_data = self.get_fields(model_class, field_name, path, '')
        context['related_fields'], model_ct, context['path'] = self.get_related_fields(model_class, field_name, path)
        context['model_ct'] = model_ct.id
        context['field_name'] = field_name
        context['table'] = True
        context = dict(context.items() + field_data.items())
        return self.render_to_response(context)
