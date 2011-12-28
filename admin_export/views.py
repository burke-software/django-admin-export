# Copyright (c) 2011, Burke Software and Consulting LLC
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# Redistributions of source code must retain the above copyright notice, this
# list of conditions and the following disclaimer.
# Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
# ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
# LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
# CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
# SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
# INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN 
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) 
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.

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
    
    if request.POST['check_default'] == "true":
        check_default = True
    else:
        check_default = False
    
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
        'many_to_many': model._meta.many_to_many,
        'previous_fields': previous_fields,
        'check_default': check_default,
    }, RequestContext(request, {}),)

def write_to_xls(worksheet, data, row_to_insert_data, ci, is_m2m):
    """ Write data in exactly one cell. For m2m comma seperate it """
    if str(data.__class__) == "<class 'django.db.models.fields.related.ManyRelatedManager'>":
        # Iterate through each m2m object and concatinate them together seperated by a comma
        m2m_data = ""
        for m2m in data.all():
            m2m_data += unicode(m2m) + ","
        data = m2m_data[:-1]
    if not is_m2m:
        if data:
            worksheet.write(row_to_insert_data, ci, unicode(data))

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
        row_to_insert_data = 0
        for ri, row in enumerate(queryset): # For Row iterable, data row in the queryset
            row_to_insert_data += 1
            added_m2m_rows = 0 # Extra rows to add to make room for many to many sub fields.
            for ci, field in enumerate(fieldnames): # For Cell iterable, field, fields
                try:
                    is_m2m = False # True if this is a sub field of a manytomany field.
                    field = field.split('__')
                    data = getattr(row, field[1])
                    for sub_field in field[2:]:
                        if str(data.__class__) == "<class 'django.db.models.fields.related.ManyRelatedManager'>":
                            is_m2m = True
                            for related_i, related_object in enumerate(data.all()):
                                data = getattr(related_object, sub_field)
                                write_to_xls(worksheet, data, row_to_insert_data+related_i, ci, False)
                                #worksheet.write(row_to_insert_data+related_i, ci, unicode(data))
                                if added_m2m_rows < related_i:
                                    added_m2m_rows = related_i
                        else:
                            data = getattr(data, sub_field)
                    write_to_xls(worksheet, data, row_to_insert_data, ci, is_m2m)
                except: # In case there is a None for a referenced field
                    pass
            row_to_insert_data += added_m2m_rows
        
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
        'many_to_many': model_class._meta.many_to_many,
        'get_variables': get_variables,
    }, RequestContext(request, {}),)
