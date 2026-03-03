import pandas as pd
import json, os, glob, datetime

files = glob.glob('data/*.xlsx') + glob.glob('data/*.xls')
if not files:
    print("No se encontró Excel en data/")
    exit(1)

excel_path = files[0]
print(f"Procesando: {excel_path}")

df = pd.read_excel(excel_path, header=1)
df = df.dropna(how='all', axis=1).dropna(how='all', axis=0)

column_map = {
    'Centro': 'Centro',
    'Rofina': 'Rofina',
    'Siegfried': 'Siegfried',
    'Descripcion': 'Descripcion',
    'Stock': 'Stock',
    'VENTAS al dia': 'Ventas/Día',
    '%': '%',
    'Estimado VTA Mazro': 'Est. VTA',
    'DIAS / ESTIMADO': 'Días/Est.',
    'VTA (prom ult 3 meses)': 'VTA Prom 3m',
    'DIAS / PROM 3': 'Días/Prom3',
    'Cuarentena / Próximos ingresos': 'Cuarentena',
    'LOTES TRANSITO': 'Lotes Tránsito',
    'OBSERVACIONES': 'Observaciones',
}
df = df.rename(columns={k: v for k, v in column_map.items() if k in df.columns})

records = df.fillna('').astype(str).to_dict(orient='records')
columns = list(df.columns)

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

print("✅ Listo: output/index.html generado")
