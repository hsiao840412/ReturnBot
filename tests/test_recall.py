import csv
import json
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
from kgb_template import template_path, N
from invoice_common import InvoiceLayout, invoice_detail

from recall import export, import_prices, prepare, xlsx_rows


def item(part='TA661-12345', quantity='1', twd='159650', description='A, original  description, '):
    return dict(part=part, quantity=quantity, twd=twd, description=description, weight='0.2', country='CN')


def payload(rows, rate='31.93'):
    return {'case': 'NAR_TEST', 'rate': rate, 'rows': rows}


class RecallTests(unittest.TestCase):
    def test_limit_quantity_and_round_before_split(self):
        result = prepare(payload([item(twd='79.84', quantity='2'), item(twd='159458.42')]))
        self.assertEqual([g['total'] for g in result['groups']], [5000])
        self.assertEqual(result['groups'][0]['rows'][0]['usd'], 3)
        self.assertEqual(result['quantity'], 3)
        result = prepare(payload([item(), item(twd='31.93')]))
        self.assertEqual([g['total'] for g in result['groups']], [5000, 1])

    def test_half_up(self):
        self.assertEqual(prepare(payload([item(twd='25')], rate='10'))['total'], 3)

    def test_invalid_or_over_limit_rows_cannot_export(self):
        for change in ({'quantity': '0'}, {'quantity': '1.5'}, {'twd': 'NaN'}, {'description': ''},
                       {'weight': '-1'}, {'country': 'China'}, {'twd': '159651', 'quantity': '2'}):
            with self.subTest(change=change), self.assertRaises(ValueError):
                prepare(payload([{**item(), **change}]))
        for rate in ('0', '-1', 'NaN', 'Infinity', ''):
            with self.subTest(rate=rate), self.assertRaises(ValueError):
                prepare(payload([item()], rate=rate))

    def test_missing_and_duplicate_prices_fail(self):
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / 'prices.csv'
            path.write_text('零件編號,交換價格,零件說明（原文）\nTA661-1,,test\n', encoding='utf-8')
            with self.assertRaises(ValueError):
                import_prices(path)
            path.write_text('零件編號,交換價格,零件說明（原文）\nTA661-1,100,test\nTA661-1,200,test\n', encoding='utf-8')
            with self.assertRaisesRegex(ValueError, '重複料號'):
                import_prices(path)

    def test_raw_price_options_currency_dedup_and_zero(self):
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / 'prices.csv'
            path.write_text('產品名稱,零件編號,零件說明,零件貨幣,價格選項,價格\n'
                            'Phone,TA661-1,"Original, text",TWD,交換價格,100.0\n'
                            'Phone 2,TA661-1,"Original, text",TWD,交換價格,100\n'
                            'Phone,TA661-1,"Original, text",TWD,庫存價格 ,300\n'
                            'Phone,TA661-1,"Original, text",USD,交換價格,9\n'
                            'Phone,TA661-2,zero,TWD,交換價格,0\n', encoding='utf-8')
            library = import_prices(path)
            self.assertEqual(len(library['entries']), 2)
            self.assertEqual(library['mergedRows'], 1)
            self.assertEqual(library['zeroPriceCount'], 1)
            self.assertEqual(library['entries'][0]['description'], 'Original, text')
            self.assertEqual(library['entries'][0]['twd'], '100.0')

    def test_zero_requires_per_line_confirmation(self):
        with self.assertRaisesRegex(ValueError, '0 美金'):
            prepare(payload([item(twd='0')]))
        result = prepare(payload([{**item(twd='0'), 'zeroPriceConfirmed': True}]))
        self.assertEqual(result['total'], 0)
        with self.assertRaises(ValueError):
            prepare(payload([{**item(twd='-1'), 'zeroPriceConfirmed': True}]))

    @unittest.skipUnless(template_path().is_file(), 'Private KGB template is not present')
    def test_export_round_trip_and_no_overwrite(self):
        description = '  Original, "quoted"  description ' + 'x' * 80 + '\nsecond line'
        with tempfile.TemporaryDirectory() as d:
            data = {**payload([item(twd='79.84', quantity='2', description=description), item()]), 'outputDirectory': d}
            result = export(data)
            folder = Path(result['outputPath'])
            csvs = sorted(folder.glob('*.csv'))
            self.assertEqual(len(csvs), 2)
            with csvs[0].open(newline='', encoding='utf-8') as f:
                rows = list(csv.reader(f))
            self.assertEqual(rows[0], ['1', 'INV_ITEM', description, '', '2', 'PCS', '3', 'USD', '0.2', '', 'CN'])
            workbook = next(folder.glob('*01*.xlsx'))
            with ZipFile(workbook) as z:
                sheet = E.fromstring(z.read('xl/worksheets/sheet1.xml'))
                self.assertEqual(sheet.find(".//m:c[@r='D13']/m:is/m:t", N).text, description)
                self.assertEqual(sheet.find(".//m:c[@r='K13']/m:v", N).text, '6')
            self.assertEqual(json.loads((folder / '召回案件.json').read_text())['rate'], '31.93')
            self.assertNotEqual(export(data)['outputPath'], result['outputPath'])

    @unittest.skipUnless(template_path().is_file(), 'Private KGB template is not present')
    def test_template_rows_formulas_print_and_stamp(self):
        with tempfile.TemporaryDirectory() as d:
            for count in (1, 3, 8, 30):
                with self.subTest(count=count):
                    data = {**payload([item(twd='31.93') for _ in range(count)]), 'outputDirectory': d}
                    result = export(data)
                    path = next(Path(result['outputPath']).glob('*.xlsx'))
                    with ZipFile(path) as z, ZipFile(template_path()) as original:
                        sheet = E.fromstring(z.read('xl/worksheets/sheet1.xml'))
                        label = E.fromstring(z.read('xl/worksheets/sheet2.xml'))
                        book = E.fromstring(z.read('xl/workbook.xml'))
                        total = sheet.find(f".//m:c[@r='K{13+count}']/m:f", N).text
                        self.assertEqual(total, f'SUM(K13:K{12+count})')
                        self.assertEqual(label.find(".//m:c[@r='B12']/m:f", N).text, f"'KBB&KGB invoice'!K{15+count}")
                        self.assertEqual(label.find(".//m:c[@r='B12']/m:v", N).text, str(count))
                        self.assertEqual(label.find(".//m:c[@r='B10']/m:v", N).text, 'NAR_TEST_01')
                        self.assertIsNone(label.find(".//m:c[@r='B15']/m:f", N))
                        area = book.find(".//m:definedName[@name='_xlnm.Print_Area'][@localSheetId='0']", N).text
                        self.assertEqual(area, f"'KBB&KGB invoice'!$A$1:$L${37+count}")
                        self.assertEqual(z.read('xl/media/image1.png'), original.read('xl/media/image1.png'))
                        self.assertNotIn('xl/calcChain.xml', z.namelist())
                        self.assertEqual(len(sheet.findall('m:mergeCells/m:mergeCell', N)), 16 + count)
                        self.assertEqual([s.get('name') for s in book.find('m:sheets', N)], ['KBB&KGB invoice', '外箱標籤'])
                        self.assertNotIn('xl/worksheets/sheet3.xml', z.namelist())
                        self.assertEqual(sheet.find('m:pageSetup', N).get('orientation'), 'portrait')
                        self.assertEqual(sheet.find("m:sheetData/m:row[@r='13']", N).get('ht'), '15.0')
                        self.assertIsNone(sheet.find(".//m:c[@r='L13']/m:v", N))
                        for xml in (sheet, label):
                            self.assertFalse(any('ePacking' in (f.text or '') for f in xml.findall('.//m:f', N)))

    def test_normal_and_recall_share_invoice_contract(self):
        layout = InvoiceLayout(8)
        self.assertEqual((layout.start, layout.delta, layout.last, layout.total_row, layout.quantity_row), (13, 5, 20, 21, 23))
        row = invoice_detail(1, 'PART', 'RMA', 'Original description', 2, 'KGB', 445, 890)
        self.assertEqual(row, [1, 'PART', 'RMA', 'Original description', None, None, None, 2, 'KGB', 445, 890, None])


if __name__ == '__main__':
    unittest.main()
