"""Consignment recall: local price import, reviewed lines, deterministic splitting.

No prices are shipped with the application. Invoice and DHL share the same
reviewed description and ROUND_HALF_UP unit price. All money uses Decimal.
"""
import csv
import io
import json
import re
import shutil
import tempfile
import zipfile
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def number(value, label, positive=True):
    try:
        result = Decimal(str(value).replace(',', '').strip())
    except (InvalidOperation, ValueError):
        raise ValueError(f'{label}必須是數字')
    if not result.is_finite() or (positive and result <= 0):
        raise ValueError(f'{label}必須大於 0')
    return result


def xlsx_rows(path):
    """Read typed/cached cell values, preserving description strings verbatim."""
    with zipfile.ZipFile(path) as z:
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            strings = [''.join(e.itertext()) for e in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        workbook = ET.fromstring(z.read('xl/workbook.xml'))
        rels = {r.attrib['Id']: r.attrib['Target'] for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        for sheet in workbook.find('m:sheets', NS):
            rid = sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            target = rels[rid]
            target = target.lstrip('/') if target.startswith('/') else 'xl/' + target
            rows = []
            for row in ET.fromstring(z.read(target)).findall('.//m:sheetData/m:row', NS):
                values = []
                for cell in row:
                    ref = re.match(r'[A-Z]+', cell.attrib['r'])[0]
                    col = 0
                    for ch in ref:
                        col = col * 26 + ord(ch) - 64
                    values.extend([''] * max(0, col - len(values)))
                    value = cell.find('m:v', NS)
                    value = value.text if value is not None and value.text else ''
                    kind = cell.attrib.get('t')
                    if kind == 's':
                        value = strings[int(value)]
                    elif kind == 'inlineStr':
                        value = ''.join(t.text or '' for t in cell.findall('.//m:t', NS))
                    values[col - 1] = value
                rows.append(values)
            yield sheet.attrib['name'], rows


def import_prices(path):
    path = Path(path)
    if path.suffix.lower() == '.csv':
        content = path.read_bytes()
        for encoding in ('utf-8-sig', 'cp950'):
            try:
                tables = [(path.stem, list(csv.reader(io.StringIO(content.decode(encoding)))))]
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError('價格表編碼無法辨識，請使用 UTF-8 CSV')
    elif path.suffix.lower() == '.xlsx':
        tables = list(xlsx_rows(path))
    else:
        raise ValueError('請匯入 .xlsx 或 .csv 價格表')
    aliases = {
        'part': ['零件編號', '零件', '料號', 'Part Number', 'Part'],
        'price': ['交換價格', '交換價', '交換單價 (TWD)', '原交換單價 (TWD)', '台幣單價', '價格', 'Price'],
        'description': ['零件說明（原文）', '零件說明', '商品描述', 'Description'],
        'option': ['價格選項'],
        'currency': ['零件貨幣', '幣別', 'Currency'],
    }
    for sheet, rows in tables:
        for header_index, header in enumerate(rows[:30]):
            names = [str(v).strip().casefold() for v in header]
            columns = {key: next((names.index(a.casefold()) for a in options if a.casefold() in names), None)
                       for key, options in aliases.items()}
            if columns['part'] is None or columns['price'] is None:
                continue
            entries = {}
            matched = 0
            source_rows = len(rows) - header_index - 1
            if columns['option'] is not None and columns['currency'] is None:
                raise ValueError('原始價格表缺少「零件貨幣」，無法確認台幣價格')
            for idx, row in enumerate(rows[header_index + 1:], header_index + 2):
                def value(key):
                    col = columns[key]
                    return row[col] if col is not None and col < len(row) else ''
                if columns['option'] is not None and str(value('option')).strip() != '交換價格':
                    continue
                if columns['currency'] is not None and str(value('currency')).strip().upper() != 'TWD':
                    continue
                part = str(value('part')).strip().upper()
                if not part:
                    continue
                # Skip footer labels; never silently discard a populated price row.
                if part in ('合計', '總計', 'TOTAL'):
                    continue
                price = number(value('price'), f'{sheet} 第 {idx} 列價格', positive=False)
                if price < 0:
                    raise ValueError(f'{sheet} 第 {idx} 列價格不可小於 0')
                entry = {'part': part, 'twd': str(price), 'description': str(value('description'))}
                existing = entries.get(part)
                if existing and (Decimal(existing['twd']) != price or existing['description'] != entry['description']):
                    raise ValueError(f'價格表重複料號 {part} 的交換價格或描述不一致，請先確認來源')
                entries.setdefault(part, entry)
                matched += 1
            if not entries:
                raise ValueError('價格表沒有商品資料')
            return {'entries': list(entries.values()), 'source': path.name, 'sheet': sheet,
                    'sourceRows': source_rows, 'matchedRows': matched,
                    'mergedRows': matched - len(entries),
                    'zeroPriceCount': sum(Decimal(e['twd']) == 0 for e in entries.values()),
                    'importedAt': datetime.now().isoformat(timespec='seconds')}
    raise ValueError('找不到料號與台幣交換價格欄位。請使用「零件編號」「交換價格」「零件說明（原文）」標題。')


def prepare(payload):
    rate = number(payload.get('rate', ''), '匯率')
    case = str(payload.get('case', '')).strip()
    if not re.fullmatch(r'[A-Za-z0-9_-]{1,60}', case):
        raise ValueError('召回單號請使用 1–60 個英文字母、數字、底線或連字號')
    rows = payload.get('rows', [])
    if not rows:
        raise ValueError('請先加入召回零件')
    groups, current, subtotal = [], [], Decimal(0)
    for index, raw in enumerate(rows, 1):
        part = str(raw.get('part', '')).strip()
        description = str(raw.get('description', ''))
        if not part or not description.strip():
            raise ValueError(f'第 {index} 筆缺少料號或商品描述')
        quantity = number(raw.get('quantity', ''), f'第 {index} 筆數量')
        if quantity != quantity.to_integral_value():
            raise ValueError(f'第 {index} 筆數量必須是整數')
        twd = number(raw.get('twd', ''), f'第 {index} 筆台幣單價', positive=False)
        if twd < 0:
            raise ValueError(f'{part} 台幣單價不可小於 0')
        usd = (twd / rate).quantize(Decimal('1'), rounding=ROUND_HALF_UP)
        if usd == 0 and raw.get('zeroPriceConfirmed') is not True:
            raise ValueError(f'{part} 美金單價為 0，請勾選該筆「確認以 0 美金申報」，或修正價格')
        weight = number(raw.get('weight', ''), f'第 {index} 筆重量')
        country = str(raw.get('country', '')).strip().upper()
        if not re.fullmatch('[A-Z]{2}', country):
            raise ValueError(f'{part} 原產地請填兩碼國別代碼')
        total = usd * quantity
        if total > 5000:
            raise ValueError(f'{part} 單筆 USD {total:,} 超過 5,000，請手動拆成多列並分配數量')
        row = {'part': part, 'description': description, 'quantity': int(quantity),
               'twd': str(twd), 'usd': int(usd), 'total': int(total),
               'weight': str(weight), 'country': country}
        if subtotal + total > 5000:
            groups.append({'rows': current, 'total': int(subtotal)})
            current, subtotal = [], Decimal(0)
        current.append(row)
        subtotal += total
    if current:
        groups.append({'rows': current, 'total': int(subtotal)})
    return {'case': case, 'rate': str(rate), 'groups': groups,
            'quantity': sum(r['quantity'] for g in groups for r in g['rows']),
            'total': sum(g['total'] for g in groups)}


def export(payload):
    from kgb_template import write_kgb, template_path
    result = prepare(payload)
    template = template_path()
    if not template.is_file():
        raise ValueError('找不到 kgb template.xlsx，請重新安裝包含 KGB 範本的版本')
    root = Path(payload['outputDirectory'])
    root.mkdir(parents=True, exist_ok=True)
    # A fresh directory on every run: no partially overwritten past shipment.
    destination = Path(tempfile.mkdtemp(prefix=result['case'] + '_', dir=root))
    try:
        for i, group in enumerate(result['groups'], 1):
            invoice = f"{result['case']}_{i:02d}"
            with (destination / f'{invoice}_DHL.csv').open('w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f, lineterminator='\r\n')
                for row in group['rows']:
                    writer.writerow([1, 'INV_ITEM', row['description'], '', row['quantity'], 'PCS', row['usd'], 'USD', row['weight'], '', row['country']])
            write_kgb(template, destination / f'{invoice}_KGB.xlsx', result, group, invoice)
        snapshot = {**payload, 'result': result, 'exportedAt': datetime.now().isoformat()}
        (destination / '召回案件.json').write_text(json.dumps(snapshot, ensure_ascii=False, indent=2), encoding='utf-8')
    except Exception:
        shutil.rmtree(destination)
        raise
    return {**result, 'outputPath': str(destination)}


def handle(operation, payload):
    if operation == 'prices':
        return import_prices(payload['path'])
    if operation == 'preview':
        return prepare(payload)
    if operation == 'export':
        return export(payload)
    raise ValueError('未知寄銷召回操作')
