from django.http import HttpResponse
from django.db.models.query import QuerySet
from openpyxl import Workbook
from StringIO import StringIO
import csv


class SpreadsheetResponseMixin(object):

    def render_excel_response(self, **kwargs):
        filename = self.get_filename(extension='xlsx')
        # Generate content
        self.data, self.headers = self.render_setup(**kwargs)
        # Setup response
        content_type = \
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        response = HttpResponse(content_type=content_type)
        response['Content-Disposition'] = \
            'attachment; filename="{0}"'.format(filename)
        # Add content and return response
        self.generate_xlsx(data=self.data, headers=self.headers, file=response)
        return response

    def render_csv_response(self, **kwargs):
        filename = self.get_filename(extension='csv')
        # Generate content
        self.data, self.headers = self.render_setup(**kwargs)
        # Build response
        content_type = 'text/csv'
        response = HttpResponse(content_type=content_type)
        response['Content-Disposition'] = \
            'attachment; filename="{0}"'.format(filename)
        # Add content to response
        self.generate_csv(data=self.data, headers=self.headers, file=response)
        return response

    def render_setup(self, **kwargs):
        # Generate content
        queryset = self.get_queryset(kwargs.get('queryset'))
        fields = self.get_fields(**kwargs)
        data = self.generate_data(queryset=queryset, fields=fields)
        headers = kwargs.get('headers')
        if not headers:
            headers = self.generate_headers(data, fields=fields)
        return data, headers

    def get_queryset(self, queryset=None):
        if queryset is None:
            try:
                queryset = self.queryset
            except AttributeError:
                raise NotImplementedError(
                    "You must provide a queryset on the class or pass it in."
                )
        return queryset

    def generate_data(self, queryset=None, fields=None):
        queryset = self.get_queryset(queryset)

        # After all that, have we got a proper queryset?
        assert isinstance(queryset, QuerySet)

        if fields:
            list_of_lists = queryset.values_list(*fields)
        else:
            list_of_lists = queryset.values_list()
        return list_of_lists

    def recursively_build_field_name(self, current_model, remaining_path):
        get_field = lambda name: current_model._meta.get_field(name)

        if '__' in remaining_path:
            foreign_key_name, path_in_related_model = remaining_path.split('__', 2)
            foreign_key_field = get_field(foreign_key_name)
            related_model = foreign_key_field.rel.to
            return [foreign_key_field.verbose_name] + \
                self.recursively_build_field_name(related_model, path_in_related_model)
        else:
            return [get_field(remaining_path).verbose_name]

    def build_field_name(self, model, path):
        name_parts = self.recursively_build_field_name(model, path)
        return ' '.join(name_parts).title()

    def generate_headers(self, data, fields=None):
        if fields is None:
            fields = self.get_fields(data.model)
        return tuple(self.build_field_name(data.model, field) for field in fields)

    def generate_xlsx(self, data, headers=None, file=None):
        wb = Workbook()
        ws = wb.get_active_sheet()

        # Put in headers
        rowoffset = 0
        if headers:
            rowoffset = 1
            for c, headerval in enumerate(headers):
                ws.cell(row=0, column=c).value = headerval

        # Put in data
        for r, row in enumerate(data):
            for c, cellval in enumerate(row):
                ws.cell(row=r + rowoffset, column=c).value = cellval
        if file:
            wb.save(file)
        return wb

    def generate_csv(self, data, headers=None, file=None):
        if not file:
            generated_csv = StringIO()
        else:
            generated_csv = file
        writer = csv.writer(generated_csv, dialect='excel')
        # Put in headers
        if headers:
            writer.writerow([unicode(s).encode('utf-8') for s in headers])

        # Put in data
        for row in data:
            writer.writerow([unicode(s).encode('utf-8') for s in row])
        return generated_csv

    def get_render_method(self, format):
        if format == 'excel':
            return self.render_excel_response
        elif format == 'csv':
            return self.render_csv_response
        raise NotImplementedError("Export format is not recognized.")

    def get_format(self, **kwargs):
        if 'format' in kwargs:
            return kwargs['format']
        elif hasattr(self, 'format'):
            return self.format
        raise NotImplementedError("Format is not defined.")

    def get_filename(self, **kwargs):
        if 'filename' in kwargs:
            return kwargs['filename']
        if hasattr(self, 'filename'):
            return self.filename
        default_filename = 'export'
        extension = kwargs.get('extension', 'out')
        return "{0}.{1}".format(default_filename, extension)

    def get_fields(self, model=None, **kwargs):
        if 'fields' in kwargs:
            return kwargs['fields']
        elif hasattr(self, 'fields') and self.fields is not None:
            return self.fields
        else:
            if hasattr(self, 'queryset') and self.queryset is not None:
                if hasattr(self.queryset, 'field_names'):
                    return self.queryset.field_names
                else:
                    model = self.queryset.model
            
            if model is None and hasattr(self, 'model') and self.model is not None:
                model = self.model
            
            if model:
                return [f.name for f in model._meta.fields]

        return ()
