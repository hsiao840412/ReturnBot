"""Fill bundled OOXML templates directly; no Excel process or automation required."""
from copy import deepcopy
from pathlib import Path
import posixpath
import re
import textwrap
import zipfile
from lxml import etree as E
from invoice_common import InvoiceLayout, invoice_detail
from kgb_template import N, tag, put, xml_bytes

REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
DRAW = {'d': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}


def col(number):
    result = ''
    while number:
        number, digit = divmod(number - 1, 26)
        result = chr(65 + digit) + result
    return result


def shift(text, delta, threshold=16):
    return re.sub(r'(\$?[A-Z]{1,3}\$?)(\d+)',
                  lambda m: m[1] + str(int(m[2]) + (delta if int(m[2]) >= threshold else 0)), text)


def normalize(sheet):
    data = sheet.find('m:sheetData', N)
    data[:] = sorted(data, key=lambda r: int(r.get('r')))
    for row in data:
        row[:] = sorted(row, key=lambda c: (len(re.match('[A-Z]+', c.get('r'))[0]), re.match('[A-Z]+', c.get('r'))[0]))
    cells = data.findall('m:row/m:c', N)
    if cells:
        max_col = max((re.match('[A-Z]+', c.get('r'))[0] for c in cells), key=lambda x: (len(x), x))
        sheet.find('m:dimension', N).set('ref', f'A1:{max_col}{max(int(r.get("r")) for r in data)}')


def defined(book, name, index, text):
    names = book.find('m:definedNames', N)
    if names is None:
        names = E.SubElement(book, tag('definedNames'))
    for old in list(names):
        if old.get('name') == name and old.get('localSheetId') == str(index):
            names.remove(old)
    E.SubElement(names, tag('definedName'), name=name, localSheetId=str(index)).text = text


def print_layout(sheet, multipage):
    props = sheet.find('m:sheetPr', N)
    if props is None:
        props = E.Element(tag('sheetPr')); sheet.insert(0, props)
    fit = props.find('m:pageSetUpPr', N)
    if fit is None:
        fit = E.SubElement(props, tag('pageSetUpPr'))
    fit.set('fitToPage', '1')
    setup = sheet.find('m:pageSetup', N)
    if setup is not None:
        setup.attrib.pop('scale', None)
        setup.set('fitToWidth', '1')
        setup.set('fitToHeight', '0' if multipage else '1')


def write_normal(template, output, df, return_type, invoice_no, date, unit_price=50):
    if len(df) < 1:
        raise ValueError('退料明細不可為空')
    with zipfile.ZipFile(template) as source:
        files = {name: source.read(name) for name in source.namelist()}
    book = E.fromstring(files['xl/workbook.xml'])
    rels = E.fromstring(files['xl/_rels/workbook.xml.rels'])
    targets = {r.get('Id'): posixpath.normpath(posixpath.join('xl', r.get('Target'))) for r in rels}
    sheets = {s.get('name'): (i, targets[s.get('{'+REL+'}id')]) for i, s in enumerate(book.find('m:sheets', N))}
    roots = {name: E.fromstring(files[path]) for name, (_, path) in sheets.items()}
    invoice = roots['KBB&KGB invoice']
    packing = roots['ePacking List']
    layout = InvoiceLayout(len(df))
    data = invoice.find('m:sheetData', N)
    prototype = deepcopy(data.find("m:row[@r='13']", N))
    merges = invoice.find('m:mergeCells', N)
    for merge in list(merges):
        if merge.get('ref') in ('D13:G13', 'D14:G14', 'D15:G15'):
            merges.remove(merge)
    for row in list(data):
        old = int(row.get('r'))
        if 13 <= old <= 15:
            data.remove(row)
        elif old >= 16:
            row.set('r', str(old + layout.delta))
            for cell in row:
                cell.set('r', shift(cell.get('r'), layout.delta))
    for element in invoice.iter():
        for key in ('ref', 'sqref'):
            if key in element.attrib:
                element.set(key, shift(element.get(key), layout.delta))
    # Keep the supplied typeface, borders and merged description area; enable wrapping.
    styles = E.fromstring(files['xl/styles.xml'])
    xfs = styles.find('m:cellXfs', N)
    wrapped = deepcopy(xfs[int(prototype.find("m:c[@r='D13']", N).get('s', '0'))])
    alignment = wrapped.find('m:alignment', N)
    if alignment is None: alignment = E.SubElement(wrapped, tag('alignment'))
    alignment.set('wrapText', '1'); wrapped.set('applyAlignment', '1')
    wrapped_id = len(xfs); xfs.append(wrapped); xfs.set('count', str(len(xfs)))
    for i, item in enumerate(df.to_dict('records'), 13):
        row = deepcopy(prototype); row.set('r', str(i))
        description = str(item.get('零件說明', ''))
        lines = sum(max(1, len(textwrap.wrap(line, width=70))) for line in description.split('\n'))
        row.set('ht', str(max(float(prototype.get('ht', '15')), lines * 15)))
        row.set('customHeight', '1')
        for cell in row: cell.set('r', re.sub(r'\d+', str(i), cell.get('r')))
        data.append(row)
        rma = str(item.get('退回訂單' if return_type in ('KBB', 'KBB Battery') else '維修', ''))
        returns = str(item.get('預期退回', 'KBB')) if return_type in ('KBB', 'KBB Battery') else 'KBB'
        values = invoice_detail(i - 12, str(item.get('零件', '')), rma, description, 1, returns, unit_price, unit_price)
        for column, value in zip('ABCDEFGHIJKL', values): put(invoice, f'{column}{i}', value)
        put(invoice, f'K{i}', unit_price, f'H{i}*J{i}')
        row.find(f"m:c[@r='D{i}']", N).set('s', str(wrapped_id))
        E.SubElement(merges, tag('mergeCell'), ref=f'D{i}:G{i}')
    merges.set('count', str(len(merges)))
    put(invoice, 'K1', invoice_no); put(invoice, 'K2', date)
    put(invoice, f'J{layout.total_row}', 'Total:')
    put(invoice, f'K{layout.total_row}', len(df)*unit_price, f'SUM(K13:K{layout.last})')
    put(invoice, f'K{layout.quantity_row}', len(df), f'SUM(H13:H{layout.last})')
    # The old templates contain prior ePacking records well beyond row 200.
    # Replace all cells, preserving styles for the new header and detail rows only.
    pdata = packing.find('m:sheetData', N)
    header = deepcopy(pdata.find("m:row[@r='1']", N))
    detail = deepcopy(pdata.find("m:row[@r='2']", N))
    pdata[:] = []
    columns = list(df.columns)
    skip = bool(columns and 'no' in str(columns[0]).lower())
    headers = columns[1:] if skip else columns
    records = df.iloc[:, 1:] if skip else df
    for number, values in enumerate([['No.'] + headers] + [[i+1]+list(values) for i, values in enumerate(records.itertuples(index=False, name=None))], 1):
        proto = header if number == 1 else detail
        row = E.Element(tag('row'), **{**(proto.attrib if proto is not None else {}), 'r': str(number)})
        pdata.append(row)
        for c, value in enumerate(values, 1):
            cell = put(packing, f'{col(c)}{number}', value)
            source_cell = proto.find(f"m:c[@r='{col(c)}{1 if number == 1 else 2}']", N) if proto is not None else None
            if source_cell is not None and source_cell.get('s'): cell.set('s', source_cell.get('s'))
    for filt in packing.findall('m:autoFilter', N): filt.set('ref', f'A1:{col(len(headers)+1)}{len(df)+1}')
    defined(book, '_xlnm.Print_Area', sheets['ePacking List'][0], f"'ePacking List'!$A$1:${col(len(headers)+1)}${len(df)+1}")
    defined(book, '_xlnm.Print_Titles', sheets['ePacking List'][0], "'ePacking List'!$1:$1")
    if '條碼' in roots and return_type == 'KBB Battery':
        barcode = roots['條碼']; bdata = barcode.find('m:sheetData', N)
        proto = deepcopy(bdata.find("m:row[@r='4']", N))
        for row in list(bdata):
            if int(row.get('r')) >= 4: bdata.remove(row)
        for i, item in enumerate(df.to_dict('records'), 4):
            row = deepcopy(proto); row.set('r', str(i))
            for cell in row: cell.set('r', re.sub(r'\d+', str(i), cell.get('r')))
            bdata.append(row)
            desc = str(item.get('零件說明', ''))
            row.set('ht', str(max(float(proto.get('ht', '40')), sum(max(1, len(textwrap.wrap(line, width=50))) for line in desc.split('\n')) * 20)))
            row.set('customHeight', '1')
            repair = str(item.get('維修', ''))
            for column, value in zip('ABCDEF', [i-3, repair, '*'+repair+'*', str(item.get('退回訂單','')), str(item.get('零件','')), str(item.get('零件說明',''))]):
                put(barcode, f'{column}{i}', value)
            put(barcode, f'C{i}', '*'+repair+'*', f'"*"&B{i}&"*"')
        barcode_style = deepcopy(xfs[int(proto.find("m:c[@r='F4']", N).get('s','0'))])
        ba = barcode_style.find('m:alignment', N)
        if ba is None: ba = E.SubElement(barcode_style, tag('alignment'))
        ba.set('wrapText', '1'); barcode_style.set('applyAlignment','1')
        barcode_style_id = len(xfs); xfs.append(barcode_style); xfs.set('count',str(len(xfs)))
        for r in range(4, len(df)+4):
            barcode.find(f".//m:c[@r='F{r}']", N).set('s',str(barcode_style_id))
        put(barcode, 'F1', len(df), f'COUNTA(A4:A{len(df)+3})')
        defined(book, '_xlnm.Print_Area', sheets['條碼'][0], f"'條碼'!$A$1:$F${len(df)+3}")
        defined(book, '_xlnm.Print_Titles', sheets['條碼'][0], "'條碼'!$1:$3")
    # Cross-sheet label references move with the invoice footer; cache their values
    # so Quick Look / readers without a calculation engine still display the result.
    for name, root in roots.items():
        if name == 'KBB&KGB invoice': continue
        for cell in root.findall('.//m:c[m:f]', N):
            formula = cell.find('m:f', N).text or ''
            if formula.startswith("'KBB&KGB invoice'!"):
                address = shift(formula.split('!', 1)[1], layout.delta)
                source = invoice.find(f".//m:c[@r='{address}']", N)
                value = source.find('m:v', N) if source is not None else None
                if address == 'K1': cached = invoice_no
                elif address == 'K2': cached = date
                elif value is not None: cached = float(value.text)
                else: cached = 0  # Match Excel's direct reference to an empty cell.
                put(root, cell.get('r'), cached, "'KBB&KGB invoice'!"+address)
    # Shift drawing anchors with the footer without rewriting images/shapes.
    bottom = 25 + layout.delta
    inv_path = sheets['KBB&KGB invoice'][1]
    rel_path = posixpath.join(posixpath.dirname(inv_path), '_rels', posixpath.basename(inv_path)+'.rels')
    if rel_path in files:
        for rel in E.fromstring(files[rel_path]):
            if rel.get('Type','').endswith('/drawing'):
                path = posixpath.normpath(posixpath.join(posixpath.dirname(inv_path), rel.get('Target')))
                drawing = E.fromstring(files[path])
                for marker in drawing.findall('.//d:row', DRAW):
                    if int(marker.text) >= 15: marker.text = str(int(marker.text)+layout.delta)
                    bottom = max(bottom, int(marker.text)+2)
                files[path] = xml_bytes(drawing)
    defined(book, '_xlnm.Print_Area', 0, f"'KBB&KGB invoice'!$A$1:$L${bottom}")
    defined(book, '_xlnm.Print_Titles', 0, "'KBB&KGB invoice'!$12:$12")
    defined(book, '_xlnm._FilterDatabase', 0, f"'KBB&KGB invoice'!$A$12:$L${layout.last}")
    for entry in list(book.find('m:definedNames', N)):
        if '#REF!' in (entry.text or '') or (entry.get('localSheetId') == str(sheets['ePacking List'][0]) and entry.get('name') == '_xlnm._FilterDatabase'):
            entry.getparent().remove(entry)
    if len(df) > 20: put(invoice, 'K3', '')
    footer = invoice.find('m:headerFooter', N)
    if footer is None:
        footer = E.Element(tag('headerFooter'))
        invoice.insert(list(invoice).index(invoice.find('m:pageSetup', N))+1, footer)
    odd = footer.find('m:oddFooter', N)
    if odd is None: odd = E.SubElement(footer, tag('oddFooter'))
    odd.text = '&CPage &P of &N'
    if return_type == 'KBB Battery':
        defined(book, '_xlnm.Print_Area', 2, "'外箱標籤  '!$A$1:$M$10")
    print_layout(invoice, len(df) > 20)
    # Remove obsolete calc chain and unused source strings (including old records).
    files.pop('xl/calcChain.xml', None)
    for path in ('xl/_rels/workbook.xml.rels', '[Content_Types].xml'):
        root = E.fromstring(files[path])
        for child in list(root):
            if 'calcChain' in str(child.attrib): root.remove(child)
        files[path] = xml_bytes(root)
    if 'xl/sharedStrings.xml' in files:
        strings = E.fromstring(files['xl/sharedStrings.xml'])
        cells = [c for root in roots.values() for c in root.findall(".//m:c[@t='s']", N)]
        used = sorted({int(c.find('m:v', N).text) for c in cells})
        mapping = {old:new for new,old in enumerate(used)}
        strings[:] = [deepcopy(strings[i]) for i in used]
        for cell in cells:
            v = cell.find('m:v', N); v.text = str(mapping[int(v.text)])
        strings.set('count', str(len(cells))); strings.set('uniqueCount', str(len(used)))
        files['xl/sharedStrings.xml'] = xml_bytes(strings)
    calc = book.find('m:calcPr', N)
    if calc is not None: calc.set('fullCalcOnLoad','1'); calc.set('forceFullCalc','1')
    for name, (_, path) in sheets.items():
        normalize(roots[name]); files[path] = xml_bytes(roots[name])
    files['xl/workbook.xml'] = xml_bytes(book); files['xl/styles.xml'] = xml_bytes(styles)
    temporary = Path(str(output)+'.tmp')
    try:
        with zipfile.ZipFile(temporary, 'w', zipfile.ZIP_DEFLATED) as target:
            for name, content in files.items(): target.writestr(name, content)
        temporary.replace(output)
    finally:
        temporary.unlink(missing_ok=True)
