import unittest
import tempfile
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
import pandas as pd
from normal_invoice import write_normal
from kgb_template import N

TEMPLATES = {'Mail in':'mail-in template.xlsx', 'Mail in Battery':'mail-in swollen template.xlsx', 'KBB':'kbb template.xlsx', 'KBB Battery':'battery kbb template.xlsx'}

def sample(count):
    return pd.DataFrame([{'No.':str(i+1), '零件':f'TA661-{10000+i}', '維修':f'000{i:04}',
                          '退回訂單':f'RMA-{i}', '零件說明':'Original,  description with trailing space ',
                          '預期退回':'KBB', '原產地':'China'} for i in range(count)])

def cell(root, address, strings):
    c = root.find(f".//m:c[@r='{address}']",N)
    if c is None:return None
    if c.get('t') == 'inlineStr':return ''.join(c.find('m:is',N).itertext())
    v = c.find('m:v',N)
    if v is None:return None
    return strings[int(v.text)] if c.get('t') == 's' else v.text

class NormalInvoiceTests(unittest.TestCase):
    def test_all_templates_row_counts_and_assets(self):
        with tempfile.TemporaryDirectory() as directory:
            for kind, template in TEMPLATES.items():
                for count in (1,3,28,205):
                    with self.subTest(kind=kind,count=count):
                        path=Path(directory)/'result.xlsx'
                        write_normal(template,path,sample(count),kind,'INVOICE-TEST','2026/09/18')
                        with ZipFile(path) as z, ZipFile(template) as original:
                            roots={name:E.fromstring(z.read(name)) for name in z.namelist() if name.startswith('xl/worksheets/sheet') and name.endswith('.xml')}
                            strings=[''.join(s.itertext()) for s in E.fromstring(z.read('xl/sharedStrings.xml'))]
                            inv=roots['xl/worksheets/sheet1.xml'];pack=roots['xl/worksheets/sheet2.xml']
                            self.assertEqual(cell(inv,'C13',strings),'RMA-0' if kind in ('KBB','KBB Battery') else '0000000')
                            self.assertEqual(cell(inv,'D13',strings),sample(1).iloc[0]['零件說明'])
                            self.assertEqual(float(cell(inv,f'K{count+13}',strings)),count*50)
                            self.assertEqual(float(cell(inv,f'K{count+15}',strings)),count)
                            self.assertEqual(len(pack.findall('m:sheetData/m:row',N)), count+1)
                            self.assertEqual(cell(pack,f'A{count+1}',strings),str(count))
                            self.assertEqual(cell(pack,'C2',strings),'0000000')
                            self.assertEqual(len([m for m in inv.findall('m:mergeCells/m:mergeCell',N) if m.get('ref').startswith('D') and 13<=int(m.get('ref').split(':')[0][1:])<=count+12]), count)
                            for root in roots.values():
                                addresses=[c.get('r') for c in root.findall('m:sheetData/m:row/m:c',N)]
                                self.assertEqual(len(addresses),len(set(addresses)))
                            for name in original.namelist():
                                if name.startswith('xl/media/'):
                                    self.assertEqual(z.read(name),original.read(name))
                            for name,root in roots.items():
                                if name=='xl/worksheets/sheet1.xml':continue
                                for c in root.findall('.//m:c[m:f]',N):
                                    f=c.find('m:f',N).text or ''
                                    if f.endswith(f'!K{count+15}'):
                                        self.assertEqual(float(cell(root,c.get('r'),strings)),count)
                            self.assertNotIn('xl/calcChain.xml',z.namelist())
                            if kind=='KBB Battery':
                                barcode=roots['xl/worksheets/sheet4.xml']
                                self.assertEqual(cell(barcode,'C4',strings),'*0000000*')
                                self.assertEqual(cell(barcode,f'D{count+3}',strings),f'RMA-{count-1}')
                                self.assertEqual(cell(barcode,'F1',strings),str(count))

    def test_empty_input_is_rejected(self):
        with self.assertRaises(ValueError):
            write_normal('kbb template.xlsx','unused.xlsx',sample(0),'KBB','TEST','2026/09/18')

    def test_generation_pipeline_never_launches_external_app(self):
        from unittest.mock import patch
        from returnbot_cli import build_worker
        with tempfile.TemporaryDirectory() as directory:
            df=sample(5);df['來源國家/地區']='China'
            csv=Path(directory)/'input.csv';df.to_csv(csv,index=False)
            with patch('pathlib.Path.home', return_value=Path(directory)), patch('subprocess.Popen', side_effect=AssertionError('Generation must not start Excel')):
                for kind in TEMPLATES:
                    worker=build_worker();worker.run_excel_task(kind,str(csv))
                    results=[]
                    while not worker.task_queue.empty(): results.append(worker.task_queue.get())
                    final=results[-1]
                    self.assertEqual(final[0:2],('result',True), final)
                    self.assertTrue(Path(final[2].splitlines()[0]).exists())
            self.assertEqual(len(list((Path(directory)/'Downloads').glob('*.xlsx'))),4)
            self.assertEqual(len(list((Path(directory)/'Downloads').glob('*.csv'))),2)
