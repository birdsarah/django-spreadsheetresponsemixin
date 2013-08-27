from django.http import HttpResponse
from django.db.models.query import QuerySet
from django.test import TestCase
from StringIO import StringIO
import mock
import pytest
import factory
from openpyxl import Workbook

from spreadsheetresponsemixin import SpreadsheetResponseMixin
from .models import MockModel


class MockModelFactory(factory.django.DjangoModelFactory):
    FACTORY_FOR = MockModel
    title = factory.Sequence(lambda n: 'title{0}'.format(n))


class GenerateDataTests(TestCase):
    def setUp(self):
        self.mixin = SpreadsheetResponseMixin()
        self.mock = MockModelFactory()
        self.mock2 = MockModelFactory()
        self.queryset = MockModel.objects.all()

    def test_assertion_error_raised_if_not_a_queryset_is_sent(self):
        with pytest.raises(AssertionError):
            self.mixin.generate_data([1])

    def test_if_queryset_is_none_gets_self_queryset(self):
        self.mixin.queryset = mock.Mock(spec=QuerySet)
        self.mixin.generate_data()
        self.mixin.queryset.values_list.assert_called_once_with()

    def test_if_no_self_queryset_raise_improperlyconfigured(self):
        with pytest.raises(NotImplementedError):
            self.mixin.generate_data()

    def test_returns_values_list_qs_if_queryset(self):
        expected_list = self.queryset.values_list()
        actual_list = self.mixin.generate_data(self.queryset)
        assert list(actual_list) == list(expected_list)

    def test_returns_values_list_if_qs_is_values_queryset(self):
        expected_list = MockModel.objects.all().values_list()
        actual_list = self.mixin.generate_data(self.queryset.values())
        assert list(actual_list) == list(expected_list)

    def test_returns_self_if_qs_is_values_list_queryset(self):
        values_list_queryset = MockModel.objects.all().values_list()
        actual_list = self.mixin.generate_data(values_list_queryset)
        assert list(actual_list) == list(values_list_queryset)

    def test_uses_specified_fields(self):
        fields = ('title',)
        values_list_queryset = MockModel.objects.all().values_list()
        # We expect it to be filtered by title,
        # even though full value_list is passed
        expected_list = MockModel.objects.all().values_list(*fields)
        actual_list = self.mixin.generate_data(values_list_queryset, fields)
        assert list(actual_list) == list(expected_list)

    def test_reverse_ordering_when_fields_specified(self):
        fields = ('title', 'id')
        actual_list = self.mixin.generate_data(self.queryset, fields)
        assert actual_list[0] == (self.mock.title, self.mock.id)


class GenerateXlsxTests(TestCase):
    def setUp(self):
        self.data = (('row1col1', 'row1col2'), ('row2col1', 'row2col2'))
        self.mixin = SpreadsheetResponseMixin()

    def _get_sheet(self, wb):
        return wb.get_active_sheet()

    def test_returns_workbook_if_no_file_passed(self):
        assert type(self.mixin.generate_xlsx(self.data)) == Workbook

    def test_if_file_is_passed_it_is_returned_with_content_added(self):
        given_content = StringIO()
        assert given_content.getvalue() == ''
        self.mixin.generate_xlsx(self.data, file=given_content)
        assert given_content.getvalue() > ''

    def test_adds_row_of_data(self):
        wb = self.mixin.generate_xlsx(self.data)
        ws = self._get_sheet(wb)
        assert ws.cell('A1').value == 'row1col1'
        assert ws.cell('B2').value == 'row2col2'

    def test_inserts_headers_if_provided(self):
        headers = ('ColA', 'ColB')
        wb = self.mixin.generate_xlsx(self.data, headers)
        ws = self._get_sheet(wb)
        assert ws.cell('A1').value == 'ColA'
        assert ws.cell('B2').value == 'row1col2'


class GenerateCsvTests(TestCase):
    def setUp(self):
        self.data = (('row1col1', 'row1col2'), ('row2col1', 'row2col2'))
        self.mixin = SpreadsheetResponseMixin()

    def test_if_no_file_passed_in_stringio_is_returned(self):
        generated_csv = self.mixin.generate_csv(self.data)
        assert type(generated_csv) == type(StringIO())

    def test_if_file_is_passed_it_is_returned_with_content_added(self):
        given_content = mock.MagicMock()
        returned_file = self.mixin.generate_csv(self.data, file=given_content)
        assert returned_file == given_content

    def test_adds_row_of_data(self):
        expected_string = \
            'row1col1,row1col2\r\nrow2col1,row2col2\r\n'
        generated_csv = self.mixin.generate_csv(self.data)
        assert generated_csv.getvalue() == expected_string

    def test_inserts_headers_if_provided(self):
        headers = ('ColA', 'ColB')
        expected_string = \
            'ColA,ColB\r\nrow1col1,row1col2\r\nrow2col1,row2col2\r\n'
        generated_csv = self.mixin.generate_csv(self.data, headers)
        assert generated_csv.getvalue() == expected_string

    def test_handles_data_with_commas_correctly_by_putting_quotes(self):
        data = (('Title is, boo', '2'), )
        expected_string = \
            '"Title is, boo",2\r\n'
        generated_csv = self.mixin.generate_csv(data)
        assert generated_csv.getvalue() == expected_string

    def test_handles_data_with_quotes_correctly(self):
        data = (('Title is, "boo"', '2'), )
        expected_string = \
            '"Title is, ""boo""",2\r\n'
        generated_csv = self.mixin.generate_csv(data)
        assert generated_csv.getvalue() == expected_string


class RenderSetupTests(TestCase):
    def setUp(self):
        self.mixin = SpreadsheetResponseMixin()
        MockModelFactory()
        self.queryset = MockModel.objects.all()
        self.mixin.queryset = self.queryset

    def test_generate_data_is_called_once_with_fields_and_queryset(self):
        self.mixin.generate_data = mock.MagicMock()
        self.mixin.render_excel_response(queryset='test', fields='testfields')
        self.mixin.generate_data.assert_called_once_with(queryset='test',
                                                         fields='testfields')

    def test_if_no_headers_passed_generate_headers_called(self):
        self.mixin.generate_headers = mock.MagicMock()
        fields = (u'title',)
        self.mixin.render_excel_response(fields=fields)
        self.mixin.generate_headers.assert_called_once_with(self.mixin.data,
                                                            fields=fields)


class RenderExcelResponseTests(TestCase):
    def setUp(self):
        self.mixin = SpreadsheetResponseMixin()
        MockModelFactory()
        self.queryset = MockModel.objects.all()
        self.mixin.queryset = self.queryset

    def test_returns_httpresponse(self):
        assert type(self.mixin.render_excel_response()) == HttpResponse

    def test_returns_xlsx_content_type(self):
        expected_content_type = \
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        response = self.mixin.render_excel_response()
        actual_content_type = response._headers['content-type'][1]
        assert actual_content_type == expected_content_type

    def test_returns_attachment_content_disposition(self):
        expected_disposition = 'attachment; filename="export.xlsx"'
        response = self.mixin.render_excel_response()
        actual_disposition = response._headers['content-disposition'][1]
        assert actual_disposition == expected_disposition

    def test_generate_xslx_is_called_with_data(self):
        self.mixin.generate_xlsx = mock.MagicMock()
        self.mixin.render_excel_response()
        data = self.mixin.generate_data(self.queryset)
        mut = self.mixin.generate_xlsx
        assert mut.call_count == 1
        assert list(mut.call_args.__getnewargs__()[0][1]['data']) == list(data)

    def test_generate_xslx_is_called_with_headers(self):
        self.mixin.generate_xlsx = mock.MagicMock()
        headers = ('ColA', 'ColB')
        self.mixin.render_excel_response(headers=headers)
        mut = self.mixin.generate_xlsx
        assert mut.call_count == 1
        assert mut.call_args.__getnewargs__()[0][1]['headers'] == headers

    def test_generate_xslx_is_called_with_response(self):
        self.mixin.generate_xlsx = mock.MagicMock()
        self.mixin.render_excel_response()
        mut = self.mixin.generate_xlsx
        assert mut.call_count == 1
        assert type(mut.call_args.__getnewargs__()[0][1]['file']) \
            == HttpResponse


class RenderCsvResponseTests(TestCase):
    def setUp(self):
        MockModelFactory()
        self.queryset = MockModel.objects.all()
        self.mixin = SpreadsheetResponseMixin()
        self.mixin.queryset = self.queryset

    def test_returns_httpresponse(self):
        assert type(self.mixin.render_csv_response()) == HttpResponse

    def test_returns_csv_content_type(self):
        expected_content_type = 'text/csv'
        response = self.mixin.render_csv_response()
        actual_content_type = response._headers['content-type'][1]
        assert actual_content_type == expected_content_type

    def test_returns_attachment_content_disposition(self):
        expected_disposition = 'attachment; filename="export.csv"'
        response = self.mixin.render_csv_response()
        actual_disposition = response._headers['content-disposition'][1]
        assert actual_disposition == expected_disposition

    def test_generate_csv_is_called_with_data(self):
        self.mixin.generate_csv = mock.MagicMock()
        self.mixin.render_csv_response()
        data = self.mixin.generate_data(self.queryset)
        mut = self.mixin.generate_csv
        assert mut.call_count == 1
        assert list(mut.call_args.__getnewargs__()[0][1]['data']) == list(data)

    def test_generate_csv_is_called_with_headers(self):
        self.mixin.generate_csv = mock.MagicMock()
        headers = ('ColA', 'ColB')
        self.mixin.render_csv_response(headers=headers)
        mut = self.mixin.generate_csv
        assert mut.call_count == 1
        assert mut.call_args.__getnewargs__()[0][1]['headers'] == headers

    def test_generate_csv_is_called_with_response(self):
        self.mixin.generate_csv = mock.MagicMock()
        self.mixin.render_csv_response()
        mut = self.mixin.generate_csv
        assert mut.call_count == 1
        assert type(mut.call_args.__getnewargs__()[0][1]['file']) \
            == HttpResponse


class GenerateHeadersTests(TestCase):

    def setUp(self):
        MockModelFactory()
        self.mixin = SpreadsheetResponseMixin()
        self.data = self.mixin.generate_data(MockModel.objects.all())

    def test_generate_headers_gets_headers_from_model_name(self):
        assert self.mixin.generate_headers(self.data) == (u'Id', u'Title')

    def test_generate_headers_only_returns_fields_if_fields_is_passed(self):
        fields = ('title',)
        assert self.mixin.generate_headers(self.data,
                                           fields=fields) == (u'Title', )
