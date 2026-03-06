import openpyxl, json, os, glob, datetime
from collections import defaultdict

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
    1:  'Rofina',
    3:  'Descripcion',
    4:  'Stock',
    5:  'Ventas/Dia',
    6:  'Pct',
    7:  'Est VTA',
    8:  'Dias Est',
    12: 'Dias Venta',
    14: 'Dias Prom3',
    15: 'Cuarentena',
    16: 'Lotes Transito',
    18: 'Observaciones',
    24: 'Gran Familia',
    25: 'Familia',
    26: 'Linea',
}
columns = list(KEEP.values())
records = []

# Build linea -> gran_familia -> [familias] structure for cascading filters
structure = defaultdict(lambda: defaultdict(set))

for row in all_rows[2:]:
    if not row or row[3] is None:
        continue
    # Saltar productos discontinuados (columna AN = índice 39)
    if len(row) > 39 and str(row[39]).strip().upper() == 'DISCONTINUADO':
        continue
    linea = str(row[26]).strip() if row[26] else '-'
    gf    = str(row[24]).strip() if row[24] else '-'
    fam   = str(row[25]).strip() if row[25] else '-'
    structure[linea][gf].add(fam)

    rec = {}
    for idx, name in KEEP.items():
        v = row[idx] if idx < len(row) else None
        if v is None:
            rec[name] = ''
        elif name == 'Pct' and isinstance(v, (int, float)):
            rec[name] = round(float(v) * 100, 1)
        elif name == 'Dias Est' and isinstance(v, (int, float)):
            rec[name] = round(float(v))
        elif name == 'Cuarentena' and isinstance(v, (int, float)):
            lotes = row[16] if len(row) > 16 and isinstance(row[16], (int, float)) else 0
            rec[name] = round(float(v) + float(lotes))
        elif name in ('Dias Venta', 'Dias Prom3') and isinstance(v, (int, float)):
            rec[name] = round(float(v))
        elif isinstance(v, float):
            rec[name] = round(v, 2)
        else:
            rec[name] = str(v).strip()
    records.append(rec)

# Serialize structure
struct_out = {l: {g: sorted(f) for g, f in gd.items()} for l, gd in sorted(structure.items())}

# Calcular total Ventas/Día directo desde Excel (antes de conversión a string)
total_vta = sum(
    row[5] for row in all_rows[2:]
    if row and row[3]
    and len(row) > 39
    and str(row[39]).strip().upper() != 'DISCONTINUADO'
    and isinstance(row[5], (int, float))
)

mtime = os.path.getmtime(excel_path)
ARG = datetime.timezone(datetime.timedelta(hours=-3))
last_update = datetime.datetime.fromtimestamp(mtime, tz=ARG).strftime('%d/%m/%Y %H:%M')

with open('index.html', 'r', encoding='utf-8') as f:
    html = f.read()

html = html.replace('__DATA__', json.dumps(records, ensure_ascii=False))
html = html.replace('__COLUMNS__', json.dumps(columns, ensure_ascii=False))
html = html.replace('__STRUCTURE__', json.dumps(struct_out, ensure_ascii=False))
html = html.replace('__LAST_UPDATE__', last_update)
html = html.replace('__TOTAL_VTA__', f"{int(total_vta):,}".replace(',', '.'))

os.makedirs('output', exist_ok=True)
with open('output/index.html', 'w', encoding='utf-8') as f:
    f.write(html)

print(f"Listo: {len(records)} registros procesados")
