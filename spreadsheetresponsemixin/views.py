from django.http import HttpResponse
from django.db.models.query import QuerySet
from openpyxl import Workbook
from StringIO import StringIO
import csv


class SpreadsheetResponseMixin(object):
    def generate_data(self, queryset=None, fields=None):
        if not queryset:
            try:
                queryset = self.queryset
            except AttributeError:
                raise NotImplementedError(
                    "You must provide a queryset on the class or pass it in."
                )
        # After all that, have we got a proper queryset?
        assert isinstance(queryset, QuerySet)

        list_of_lists = queryset.values_list()
        if fields:
            list_of_lists = queryset.values_list(*fields)

        # Process into a list of lists ordered by the keys
        return list_of_lists

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
            writer.writerow(headers)

        # Put in data
        for row in data:
            writer.writerow(row)
        return generated_csv

    def generate_headers(self, data, fields=None):
        model_fields = (field for field in data.model._meta.fields)
        if fields:
            model_fields = (field
                            for field in model_fields
                            if field.name in fields)
        field_names = (field.verbose_name.title() for field in model_fields)
        return tuple(field_names)

    def render_excel_response(self, **kwargs):
        # Generate content
        self.data, self.headers = self.render_setup(**kwargs)
        # Setup response
        content_type = \
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        response = HttpResponse(content_type=content_type)
        response['Content-Disposition'] = 'attachment; filename="export.xlsx"'
        # Add content and return response
        self.generate_xlsx(data=self.data, headers=self.headers, file=response)
        return response

    def render_csv_response(self, **kwargs):
        # Generate content
        self.data, self.headers = self.render_setup(**kwargs)
        # Build response
        content_type = 'text/csv'
        response = HttpResponse(content_type=content_type)
        response['Content-Disposition'] = 'attachment; filename="export.csv"'
        # Add content to response
        self.generate_csv(data=self.data, headers=self.headers, file=response)
        return response

    def render_setup(self, **kwargs):
        # Generate content
        queryset = kwargs.get('queryset')
        fields = kwargs.get('fields')
        data = self.generate_data(queryset=queryset, fields=fields)
        headers = kwargs.get('headers')
        if not headers:
            headers = self.generate_headers(data, fields=fields)
        return data, headers
