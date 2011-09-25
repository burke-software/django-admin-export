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

from django.contrib.contenttypes.models import ContentType
from django.http import HttpResponse, HttpResponseRedirect
from django.db.models import get_model
from django.shortcuts import render_to_response, get_object_or_404
from django.template import RequestContext

import tempfile
import os
import xlwt

def get_fields_for_model(request):
    """ Get the related fields of a selected foreign key """
    model_class = ContentType.objects.get(id=request.GET['ct']).model_class()
    queryset = model_class.objects.filter(pk__in=request.GET['ids'].split(','))
    
    rel_name = request.POST['rel_name']
    related = model_class
    for item in rel_name.split('__'):
        related = getattr(related, item).field.rel.to
    
    model = related
    model_fields = model._meta.fields
    previous_fields = rel_name
    
    for field in model_fields:
        if hasattr(field, 'related'):
            if request.user.has_perm(field.rel.to._meta.app_label + '.view_' + field.rel.to._meta.module_name)\
            or request.user.has_perm(field.rel.to._meta.app_label + '.change_' + field.rel.to._meta.module_name):
                field.perm = True
    
    return render_to_response('admin_export/export_to_xls_related.html', {
        'model_name': model_class._meta.verbose_name,
        'model': model._meta.app_label + ":" + model._meta.module_name,
        'fields': model_fields,
        'previous_fields': previous_fields,
    }, RequestContext(request, {}),)

def admin_export_xls(request):
    model_class = ContentType.objects.get(id=request.GET['ct']).model_class()
    queryset = model_class.objects.filter(pk__in=request.GET['ids'].split(','))
    get_variables = request.META['QUERY_STRING']
    model_fields = model_class._meta.fields
    
    for field in model_fields:
        if hasattr(field, 'related'):
            if request.user.has_perm(field.rel.to._meta.app_label + '.view_' + field.rel.to._meta.module_name)\
            or request.user.has_perm(field.rel.to._meta.app_label + '.change_' + field.rel.to._meta.module_name):
                field.perm = True
    
    if 'xls' in request.POST:
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet(unicode(model_class._meta.verbose_name_plural))
        
        # Get field names from POST data
        fieldnames = []
        # request.POST reorders the data :( There's little reason to go through all
        # the work of reordering it right again when raw data is ordered correctly.
        for value in request.raw_post_data.split('&'):
            if value[:7] == "field__" and value[-3:] == "=on":
                fieldname = value[7:-3]
                app = fieldname.split('__')[0].split('%3A')[0]
                model = fieldname.split('__')[0].split('%3A')[1]
                # Server side permission check, edit implies view.
                if request.user.has_perm(app + '.view_' + model) or request.user.has_perm(app + '.change_' + model):
                    fieldnames.append(fieldname)
                
        # Title
        for i, field in enumerate(fieldnames):
            #ex field 'sis%3Astudent__fname'
            field = field.split('__')
            model = get_model(field[0].split('%3A')[0], field[0].split('%3A')[1])
            txt = ""
            for sub_field in field[1:-1]:
                txt += sub_field + " "
            txt += unicode(model._meta.get_field_by_name(field[-1])[0].verbose_name)
            worksheet.write(0,i, txt)
        
        # Data
        for ri, row in enumerate(queryset): # For Row iterable, data row in the queryset
            for ci, field in enumerate(fieldnames): # For Cell iterable, field, fields
                try:
                    field = field.split('__')
                    data = getattr(row, field[1])
                    for sub_field in field[2:]:
                        data = getattr(data, sub_field)
                    worksheet.write(ri+1, ci, unicode(data))
                except: # In case there is a None for a referenced field
                    pass 
        
        # Boring file handeling crap
        fd, fn = tempfile.mkstemp()
        os.close(fd)
        workbook.save(fn)
        fh = open(fn, 'rb')
        resp = fh.read()
        fh.close()
        response = HttpResponse(resp, mimetype='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="%s.xls"' % \
              (unicode(model_class._meta.verbose_name_plural),)
        return response
    
    return render_to_response('admin_export/export_to_xls.html', {
        'model_name': model_class._meta.verbose_name,
        'model': model_class._meta.app_label + ":" +  model_class._meta.module_name,
        'fields': model_fields,
        'get_variables': get_variables,
    }, RequestContext(request, {}),)
