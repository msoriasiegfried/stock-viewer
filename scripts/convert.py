import openpyxl, json, os, glob, datetime

files = glob.glob('data/*.xlsx') + glob.glob('data/*.xls') + glob.glob('data/*.xlsm')
if not files:
    print("No Excel found in data/")
    exit(1)

excel_path = files[0]
print(f"Processing: {excel_path}")

wb = openpyxl.load_workbook(excel_path, read_only=True, keep_vba=True, data_only=True)
ws = wb['Resumen']
all_rows = list(ws.iter_rows(values_only=True))
wb.close()

KEEP = {
    0: 'Centro', 1: 'Rofina', 2: 'Siegfried', 3: 'Descripcion',
    4: 'Stock', 5: 'Ventas/Dia', 6: 'Pct', 7: 'Est VTA',
    8: 'Dias Est', 9: 'VTA Prom 3m', 14: 'Dias Prom3',
    15: 'Cuarentena', 16: 'Lotes Transito', 17: 'Solicitud',
    18: 'Observaciones', 24: 'Gran Familia', 25: 'Familia', 26: 'Linea'
}
columns = list(KEEP.values())
records = []

for row in all_rows[2:]:
    if not row or row[3] is None:
        continue
    rec = {}
    for idx, name in KEEP.items():
        v = row[idx] if idx < len(row) else None
        if v is None:
            rec[name] = ''
        elif isinstance(v, float):
            rec[name] = round(v, 2)
        else:
            rec[name] = str(v).strip()
    records.append(rec)

mtime = os.path.getmtime(excel_path)
last_update = datetime.datetime.fromtimestamp(mtime).strftime('%d/%m/%Y %H:%M')

with open('index.html', 'r', encoding='utf-8') as f:
    html = f.read()

html = html.replace('__DATA__', json.dumps(records, ensure_ascii=False))
html = html.replace('__COLUMNS__', json.dumps(columns, ensure_ascii=False))
html = html.replace('__LAST_UPDATE__', last_update)

os.makedirs('output', exist_ok=True)
with open('output/index.html', 'w', encoding='utf-8') as f:
    f.write(html)

print(f"Listo: {len(records)} registros procesados")
