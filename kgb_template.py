"""Generate KGB invoices with the normal KBB layout and no ePacking sheet."""
from copy import deepcopy
from datetime import datetime
from pathlib import Path
import re
import zipfile
import textwrap
from invoice_common import InvoiceLayout, invoice_detail

from lxml import etree as E

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
N = {'m': MAIN}


def template_path():
    return Path(__file__).resolve().parent / 'kgb template.xlsx'


def tag(name):
    return '{' + MAIN + '}' + name


def shift_refs(text, delta):
    return re.sub(r'(\$?[A-Z]{1,3}\$?)(\d+)',
                  lambda m: m[1] + str(int(m[2]) + (delta if int(m[2]) >= 16 else 0)), text)


def put(sheet, address, value=None, formula=None):
    data = sheet.find('m:sheetData', N)
    row_no = int(re.search(r'\d+', address)[0])
    row = data.find(f"m:row[@r='{row_no}']", N)
    if row is None:
        row = E.SubElement(data, tag('row'), r=str(row_no))
    cell = row.find(f"m:c[@r='{address}']", N)
    if cell is None:
        cell = E.SubElement(row, tag('c'), r=address)
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop('t', None)
    if formula is not None:
        if isinstance(value, str):
            cell.set('t', 'str')
        E.SubElement(cell, tag('f')).text = formula
        if value is not None:
            E.SubElement(cell, tag('v')).text = str(value)
    elif value is None:
        pass
    elif isinstance(value, (int, float)):
        E.SubElement(cell, tag('v')).text = str(value)
    else:
        cell.set('t', 'inlineStr')
        text = E.SubElement(E.SubElement(cell, tag('is')), tag('t'))
        text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        text.text = str(value)
    return cell


def xml_bytes(root):
    return E.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def kbb_layout_files(template):
    """Use the normal KBB document, removing its ePacking worksheet entirely."""
    with zipfile.ZipFile(Path(template).with_name('kbb template.xlsx')) as z:
        files = {name: z.read(name) for name in z.namelist()}
    book = E.fromstring(files['xl/workbook.xml'])
    sheets = book.find('m:sheets', N)
    if [s.get('name') for s in sheets] != ['KBB&KGB invoice', 'ePacking List', '外箱標籤']:
        raise ValueError('一般 KBB 範本工作表格式不符')
    sheets.remove(sheets[1])
    for defined in list(book.find('m:definedNames', N)):
        if defined.get('localSheetId') == '1' or '#REF!' in (defined.text or ''):
            defined.getparent().remove(defined)
        elif defined.get('localSheetId') == '2':
            defined.set('localSheetId', '1')
    files['xl/workbook.xml'] = xml_bytes(book)
    files['xl/worksheets/sheet2.xml'] = files.pop('xl/worksheets/sheet3.xml')
    files.pop('xl/worksheets/_rels/sheet2.xml.rels', None)
    if 'xl/worksheets/_rels/sheet3.xml.rels' in files:
        files['xl/worksheets/_rels/sheet2.xml.rels'] = files.pop('xl/worksheets/_rels/sheet3.xml.rels')
    rels = E.fromstring(files['xl/_rels/workbook.xml.rels'])
    for rel in list(rels):
        if rel.get('Target') == 'worksheets/sheet2.xml':
            rels.remove(rel)
        elif rel.get('Target') == 'worksheets/sheet3.xml':
            rel.set('Target', 'worksheets/sheet2.xml')
    files['xl/_rels/workbook.xml.rels'] = xml_bytes(rels)
    types = E.fromstring(files['[Content_Types].xml'])
    for entry in list(types):
        if entry.get('PartName') == '/xl/worksheets/sheet3.xml':
            types.remove(entry)
    files['[Content_Types].xml'] = xml_bytes(types)
    # Keep the shipper/importer information and stamp supplied in the KGB file.
    # Both templates currently have identical addresses/stamp; copy explicitly.
    with zipfile.ZipFile(template) as z:
        supplied = E.fromstring(z.read('xl/worksheets/sheet1.xml'))
        shared = E.fromstring(z.read('xl/sharedStrings.xml'))
        source_strings = [''.join(t.itertext()) for t in shared]
        invoice = E.fromstring(files['xl/worksheets/sheet1.xml'])
        for address in ('B1', 'B2', 'B3', 'B7', 'H7'):
            cell = supplied.find(f".//m:c[@r='{address}']", N)
            if cell is not None:
                value = cell.find('m:v', N)
                if value is not None:
                    text = source_strings[int(value.text)] if cell.get('t') == 's' else value.text
                    put(invoice, address, text)
        files['xl/worksheets/sheet1.xml'] = xml_bytes(invoice)
        files['xl/media/image1.png'] = z.read('xl/media/image1.png')
    return files


def write_kgb(template, output, result, group, invoice):
    files = kbb_layout_files(template)
    book = E.fromstring(files['xl/workbook.xml'])
    names = [s.get('name') for s in book.find('m:sheets', N)]
    if names != ['KBB&KGB invoice', '外箱標籤']:
        raise ValueError('KGB 範本工作表格式不符，請使用提供的範本')
    sheet = E.fromstring(files['xl/worksheets/sheet1.xml'])
    label = E.fromstring(files['xl/worksheets/sheet2.xml'])
    data = sheet.find('m:sheetData', N)
    prototype = data.find("m:row[@r='13']", N)
    if prototype is None or data.find("m:row[@r='16']/m:c[@r='K16']/m:f", N) is None:
        raise ValueError('KGB 範本的明細列或合計位置不符')
    prototype = deepcopy(prototype)
    count = len(group['rows'])
    layout = InvoiceLayout(count)
    delta, last = layout.delta, layout.last
    total_row, qty_row = layout.total_row, layout.quantity_row
    merges = sheet.find('m:mergeCells', N)
    for merge in list(merges):
        if merge.get('ref') in ('D13:G13', 'D14:G14', 'D15:G15'):
            merges.remove(merge)
    for row in list(data):
        old = int(row.get('r'))
        if 13 <= old <= 15:
            data.remove(row)
        elif old >= 16:
            row.set('r', str(old + delta))
            for cell in row:
                cell.set('r', shift_refs(cell.get('r'), delta))
                formula = cell.find('m:f', N)
                if formula is not None and formula.text:
                    formula.text = shift_refs(formula.text, delta)
    for element in sheet.iter():
        for key in ('ref', 'sqref'):
            if key in element.attrib:
                element.set(key, shift_refs(element.get(key), delta))
        if E.QName(element).localname == 'formula' and element.text:
            element.text = shift_refs(element.text, delta)
    styles = E.fromstring(files['xl/styles.xml'])
    xfs = styles.find('m:cellXfs', N)
    description_style = deepcopy(xfs[int(prototype.find("m:c[@r='D13']", N).get('s'))])
    alignment = description_style.find('m:alignment', N)
    if alignment is None:
        alignment = E.SubElement(description_style, tag('alignment'))
    alignment.set('wrapText', '1')
    description_style.set('applyAlignment', '1')
    style_id = len(xfs)
    xfs.append(description_style)
    xfs.set('count', str(len(xfs)))
    merges = sheet.find('m:mergeCells', N)
    for i, item in enumerate(group['rows'], 13):
        row = deepcopy(prototype)
        row.set('r', str(i))
        lines = sum(max(1, len(textwrap.wrap(line, width=78))) for line in item['description'].split('\n'))
        row.set('ht', str(max(float(prototype.get('ht', '15')), lines * 15)))
        row.set('customHeight', '1')
        for cell in row:
            cell.set('r', re.sub(r'\d+', str(i), cell.get('r')))
        data.append(row)
        values = invoice_detail(i - 12, item['part'], result['case'], item['description'],
                                item['quantity'], 'KGB', item['usd'], item['total'])
        for col, value in zip('ABCDEFGHIJKL', values):
            put(sheet, f'{col}{i}', value)
        put(sheet, f'K{i}', item['total'], f'H{i}*J{i}')
        sheet.find(f"m:sheetData/m:row[@r='{i}']/m:c[@r='D{i}']", N).set('s', str(style_id))
        E.SubElement(merges, tag('mergeCell'), ref=f'D{i}:G{i}')
    merges.set('count', str(len(merges)))
    data[:] = sorted(data, key=lambda r: int(r.get('r')))
    for row in data:
        row[:] = sorted(row, key=lambda c: (len(re.match('[A-Z]+', c.get('r'))[0]), re.match('[A-Z]+', c.get('r'))[0]))
    date = datetime.now().strftime('%Y/%m/%d')
    qty = sum(r['quantity'] for r in group['rows'])
    put(sheet, 'K1', invoice)
    put(sheet, 'K2', date)
    if count > 20:
        put(sheet, 'K3', '')
    put(sheet, f'K{total_row}', group['total'], f'SUM(K13:K{last})')
    put(sheet, f'K{qty_row}', qty, f'SUM(H13:H{last})')
    put(sheet, f'J{qty_row}', None)  # Actual packed gross weight, not item net weight.
    put(sheet, f'L{qty_row}', None)  # Actual carton dimensions.
    put(label, 'B10', invoice, "'KBB&KGB invoice'!K1")
    put(label, 'A1', 'KGB')
    put(label, 'B11', '1/1')
    put(label, 'B12', qty, f"'KBB&KGB invoice'!K{qty_row}")
    put(label, 'B13', '', f'IF(\'KBB&KGB invoice\'!L{qty_row}="","",\'KBB&KGB invoice\'!L{qty_row})')
    put(label, 'B16', date, "'KBB&KGB invoice'!K2")
    for defined in book.findall('m:definedNames/m:definedName', N):
        if defined.get('localSheetId') == '0':
            if defined.get('name') == '_xlnm.Print_Area':
                defined.text = f"'KBB&KGB invoice'!$A$1:$L${40 + delta}"
            elif defined.get('name') == '_xlnm._FilterDatabase':
                defined.text = f"'KBB&KGB invoice'!$A$12:$L${last}"
    defined_names = book.find('m:definedNames', N)
    # Header rows repeat if the invoice spans multiple printed pages.
    E.SubElement(defined_names, tag('definedName'), name='_xlnm.Print_Titles', localSheetId='0').text = "'KBB&KGB invoice'!$12:$12"
    setup = sheet.find('m:pageSetup', N)
    setup.attrib.pop('scale', None)
    setup.set('fitToWidth', '1')
    setup.set('fitToHeight', '1' if count <= 20 else '0')
    properties = sheet.find('m:sheetPr', N)
    if properties is None:
        properties = E.Element(tag('sheetPr'))
        sheet.insert(0, properties)
    fit = properties.find('m:pageSetUpPr', N)
    if fit is None:
        fit = E.SubElement(properties, tag('pageSetUpPr'))
    fit.set('fitToPage', '1')
    footer = sheet.find('m:headerFooter', N)
    if footer is None:
        footer = E.Element(tag('headerFooter'))
        sheet.insert(list(sheet).index(setup) + 1, footer)
    E.SubElement(footer, tag('oddFooter')).text = '&CPage &P of &N'
    calc = book.find('m:calcPr', N)
    calc.set('fullCalcOnLoad', '1')
    calc.set('forceFullCalc', '1')
    # Drop obsolete calculation chain; Excel reconstructs it on opening.
    files.pop('xl/calcChain.xml', None)
    for name in ('xl/_rels/workbook.xml.rels', '[Content_Types].xml'):
        root = E.fromstring(files[name])
        for child in list(root):
            if 'calcChain' in str(child.attrib):
                root.remove(child)
        files[name] = xml_bytes(root)
    files['xl/workbook.xml'] = xml_bytes(book)
    files['xl/worksheets/sheet1.xml'] = xml_bytes(sheet)
    files['xl/worksheets/sheet2.xml'] = xml_bytes(label)
    files['xl/styles.xml'] = xml_bytes(styles)
    # The supplied stamp image is anchored below the remarks, not a header logo.
    # Move both anchors with the footer and include the entire stamp in printing.
    drawing = E.fromstring(files['xl/drawings/drawing1.xml'])
    drawing_ns = {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
    for marker in drawing.findall('.//xdr:row', drawing_ns):
        if int(marker.text) >= 15:
            marker.text = str(int(marker.text) + delta)
    files['xl/drawings/drawing1.xml'] = xml_bytes(drawing)
    # Remove strings that belonged only to the removed ePacking worksheet.
    strings = E.fromstring(files['xl/sharedStrings.xml'])
    used = sorted({int(c.find('m:v', N).text) for root in (sheet, label)
                   for c in root.findall(".//m:c[@t='s']", N)})
    mapping = {old: new for new, old in enumerate(used)}
    kept = [deepcopy(strings[index]) for index in used]
    references = 0
    for root in (sheet, label):
        for cell in root.findall(".//m:c[@t='s']", N):
            value = cell.find('m:v', N)
            value.text = str(mapping[int(value.text)])
            references += 1
    strings[:] = kept
    strings.set('count', str(references))
    strings.set('uniqueCount', str(len(kept)))
    files['xl/sharedStrings.xml'] = xml_bytes(strings)
    files['xl/worksheets/sheet1.xml'] = xml_bytes(sheet)
    files['xl/worksheets/sheet2.xml'] = xml_bytes(label)
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, content in files.items():
            z.writestr(name, content)
