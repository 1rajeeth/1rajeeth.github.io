import csv
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

src = Path('/Users/rajeeth/Downloads/Virtual Wardrobe & Care Master Tracker.xlsx')
dest = Path('work/notion-import')
dest.mkdir(parents=True, exist_ok=True)
ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
z = zipfile.ZipFile(src)
strings = [''.join(t.text or '' for t in si.findall('.//m:t', ns)) for si in ET.fromstring(z.read('xl/sharedStrings.xml'))]
book = ET.fromstring(z.read('xl/workbook.xml'))
rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
targets = {r.attrib['Id']: r.attrib['Target'] for r in rels}
rid = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
wanted = {'Wardrobe & Shoes': 'Wardrobe.csv', 'Grooming & Fragrances': 'Grooming Inventory.csv'}
for sheet in book.findall('.//m:sheet', ns):
    name = sheet.attrib['name']
    if name not in wanted:
        continue
    target = targets[sheet.attrib[rid]].lstrip('/')
    path = target if target.startswith('xl/') else 'xl/' + target
    root = ET.fromstring(z.read(path))
    rows = []
    for row in root.findall('.//m:sheetData/m:row', ns):
        cells = {}
        for cell in row.findall('m:c', ns):
            ref = cell.attrib['r']
            col = 0
            for ch in ref.rstrip('0123456789'):
                col = col * 26 + ord(ch) - 64
            value = cell.find('m:v', ns)
            value = value.text if value is not None else ''
            if cell.attrib.get('t') == 's' and value:
                value = strings[int(value)]
            cells[col - 1] = value
        if cells:
            rows.append([cells.get(i, '') for i in range(max(cells) + 1)])
    with (dest / wanted[name]).open('w', newline='', encoding='utf-8-sig') as f:
        csv.writer(f).writerows(rows)
    print(f'{name}: {len(rows)-1} rows')
