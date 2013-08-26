django-spreadsheetresponsemixin
===============================

View mixin for django, that generates a csv or excel sheet.

Usage
=====

Add to your django view as a mixin::

    class ExcelExportView(SpreadsheetResponseMixin, ListView):
        def get(self, request):
            self.queryset = self.get_queryset()
            return self.render_excel_response()


    class CsvExportView(SpreadsheetResponseMixin, ListView):
        def get(self, request):
            self.queryset = self.get_queryset()
            return self.render_csv_response()

Note you must specify a Queryset, ValuesQueryset or ValuesListQueryset on the 
class or pass it in when you call the render method.

You can also specify the fields and the headers as tuples if you want to refine
the results and / or provide custom headers for your columns
