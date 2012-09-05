# Copyright (c) 2011-2012, Burke Software and Consulting LLC
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

import django
from django.contrib.contenttypes.models import ContentType
from django.http import HttpResponse, HttpResponseRedirect
from django.db.models import get_model
from django.shortcuts import render_to_response, get_object_or_404
from django.template import RequestContext
from django.utils.encoding import smart_unicode

import tempfile
import os
import xlwt

def get_fields_for_model(request, initial=False):
    model_class = ContentType.objects.get(id=request.GET['ct']).model_class()
    queryset = model_class.objects.filter(pk__in=request.GET['ids'].split(','))
    get_variables = request.META['QUERY_STRING']
    model_fields = []
    
    previous_fields = None
    check_default = False
    if request.POST:
        if request.POST['check_default'] == "true":
            check_default = True
        if request.POST['rel_name']:
            previous_fields = request.POST['rel_name']
            for item in request.POST['rel_name'].split('__'):
                if hasattr(model_class, item) and hasattr(getattr(model_class, item), 'related'):
                    model_class = getattr(model_class, item).related.model
                else:
                    field, field_model, direct, m2m  = model_class._meta.get_field_by_name(item)
                    if direct:
                        model_class = getattr(model_class, item).field.rel.to
                    else:
                        model_class = field.model
    
    for field_name in model_class._meta.get_all_field_names():
        field, field_model, direct, m2m = model_class._meta.get_field_by_name(field_name)
        field.direct = direct
        field.m2m = m2m
        
        if not direct and not m2m and not hasattr(field,'related'):
            field.related = True
        
        if not direct and not m2m:
            # These are fields outside the model that relate to this model
            # There is no easy way to get the names, so we just set them here
            field.name = field.get_accessor_name()
            field.verbose_name = field.model._meta.verbose_name
            model_fields += [field]
        elif hasattr(field, 'related'):
            if request.user.has_perm(field.rel.to._meta.app_label + '.view_' + field.rel.to._meta.module_name)\
            or request.user.has_perm(field.rel.to._meta.app_label + '.change_' + field.rel.to._meta.module_name):
                model_fields += [field]
        elif direct and not m2m:
            model_fields += [field]
            
    # Sort fields so FK and M2M go to bottom - this makes the reports MUCH neater keeping related stay stuff together
    m2m_fields = []
    direct_fields = []
    related_fields = []
    for field in model_fields:
        if field.m2m:
            m2m_fields += [field]
        elif field.direct:
            direct_fields += [field]
        else:
            related_fields += [field]
    
    fields = direct_fields + related_fields + m2m_fields
    
    
    if initial:
        template = 'admin_export/export_to_xls.html'
    else:
        template = 'admin_export/fields.html'
    return render_to_response(template, {
        'model_name': model_class._meta.verbose_name,
        'fields': fields,
        'previous_fields': previous_fields,
        'check_default': check_default,
        'get_variables':get_variables,
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
        if data and ci < 256:
            worksheet.write(row_to_insert_data, ci, smart_unicode(data))

def name_to_title(model, name):
    """
    Determines the title for a name
    name example: placement__cras__fname'
    where placement is a fk of instance, cra's is a m2m, and fname is just a field
    Essentially, it just replaces the field names with their verbose names
    result: placement cras First name
    """
    result = ""
    for part in name.split('__'):
        if hasattr(model, part) and hasattr(getattr(model, part), 'related'):
            field = getattr(model, part).related
            direct = False
            m2m = False
            field_model = None
        else:
            field, field_model, direct, m2m = model._meta.get_field_by_name(part)
        
        # Add the verbose name to the title
        if direct:
            result += field.verbose_name + " "
        else:
            result += field.model()._meta.verbose_name + " "
        
        # If we moved models, we need change to the next model
        if m2m:
            if field.__class__.__name__ == 'RelatedObject':
                model = field.field.model
            else:
                model = getattr(model,part).field.related.parent_model
        elif direct and field.rel:
            model = getattr(model,part).field.related.parent_model
        elif field.__class__ == django.db.models.related.RelatedObject:
            model = field.model
            
    return result.strip()


def name_to_data(instance, name):
    """
    With a instance and a long name this function determines it's value
    name example: placement__cras__fname
    where placement is a fk, cra's is a m2m, and fname is just a field
    If the field is a reference it will return the reference (such as a FK)
    It's you to you to deal with this, it won't just give you a string.
    Returns 2 items:
        either field or array of m2m fieldsex: 'David' or ["David", "Bob"]
        number of m2m rows
    It cannot handle more than one m2m!
    """
    result = instance # we will eventually return this as the answer
    m2m_result = []
    m2m_count = 0
    # start by dividing the name by the __ which are seperators
    for i, part in enumerate(name.split('__')):
        # check if it exists
        try:
            field, field_model, direct, m2m = result._meta.get_field_by_name(part)
            if not direct and not m2m and field.__class__ == django.db.models.related.RelatedObject:
                part = field.get_accessor_name()
        except:
            pass
        if hasattr(result, part):
            result = getattr(result, part)
            if result.__class__.__name__ in ['ManyRelatedManager','RelatedManager'] and part != name.split('__')[-1]:
                for m2m_object in result.all():
                    m2m_result += [name_to_data(m2m_object, '__'.join(name.split('__')[i+1:]))[0]]
                    m2m_count += 1
            if hasattr(result,'all'):
                result = result.all()
            
    if m2m_result:
        return m2m_result, m2m_count
    return result, m2m_count

def admin_export_xls(request):
    model_class = ContentType.objects.get(id=request.GET['ct']).model_class()
    queryset = model_class.objects.filter(pk__in=request.GET['ids'].split(','))
    get_variables = request.META['QUERY_STRING']
    
    if 'xls' in request.POST:
        workbook = xlwt.Workbook()
        # Remove : which isn't valid in a xls sheet name
        worksheet = workbook.add_sheet(smart_unicode(model_class._meta.verbose_name_plural).replace(':','')[:30])
        
        # Get field names from POST data
        fieldnames = []
        # request.POST reorders the data :( There's little reason to go through all
        # the work of reordering it right again when raw data is ordered correctly.
        for value in request.raw_post_data.split('&'):
            if value[:7] == "field__" and value[-3:] == "=on":
                fieldname = value[7:-3]
                fieldnames.append(fieldname)
                
        # Title
        for i, field in enumerate(fieldnames):
            if i < 256: # xls format limit!
                txt=name_to_title(model_class, field)
                worksheet.write(0,i, txt)
        
        # Data
        row_to_insert_data = 1
        for row in queryset: # For Row iterable, data row in the queryset
            added_rows = 1 # Extra rows to add to make room for many to many sub fields.
            for ci, field in enumerate(fieldnames): # For Cell iterable, field, fields
                    #try:
                    data, m2m_count = name_to_data(row, field)
                    if m2m_count:
                        # Adding multiple fors for one item in original queryset
                        if m2m_count > added_rows:
                            added_rows = m2m_count
                        for mi, m2m_field in enumerate(data):
                            write_to_xls(worksheet, m2m_field, row_to_insert_data + mi, ci, False)
                    else: # Simple add one cell of data
                        write_to_xls(worksheet, data, row_to_insert_data, ci, False)
                    #except:
                    #    pass
            row_to_insert_data += added_rows
        
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
    else:
        return get_fields_for_model(request, initial=True)
