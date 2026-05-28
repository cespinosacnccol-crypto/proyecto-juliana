import re, openpyxl, pandas as pd

wb = openpyxl.load_workbook('INFORME CLIENTE/ACUMULADO GENERAL.xlsx')
dfs = []
for sheet in wb.sheetnames:
    ws = wb[sheet]
    headers = [str(c.value) if c.value else f'COL{i}' for i, c in enumerate(ws[1])]
    eval_cols = [i for i, h in enumerate(headers) if h.endswith('_EVAL')]
    meta_nombres = ['NOMBRE SEDE', 'CODIGO DANE SEDE', 'NOMBRES ESTUDIANTE',
                     'COD. EST.', 'GRADO', 'CURSO', 'PRUEBA']
    meta_idx = [(i, headers.index(c)) for i, c in enumerate(meta_nombres) if c in headers]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(c is None or str(c).strip() == '' for c in row): continue
        correctas = sum(1 for i in eval_cols if i < len(row) and str(row[i]).strip().upper() == 'CORRECTO')
        total = len(eval_cols)
        meta = {}
        for mi, col_i in meta_idx:
            v = row[col_i] if col_i < len(row) and row[col_i] is not None else ''
            val = re.sub(r'\s+', ' ', str(v).strip().upper())
            meta[meta_nombres[mi]] = val
        meta['_KEY'] = meta.get('COD. EST.', '') or meta.get('NOMBRES ESTUDIANTE', '')
        meta['_correctas'] = correctas
        meta['_total'] = total
        rows.append(meta)
    if rows: dfs.append(pd.DataFrame(rows))
wb.close()
df = pd.concat(dfs, ignore_index=True)

# Fix: replace actual special chars in column names
df.columns = [c.replace('Í', 'I').replace('Ó', 'O').replace('É', 'E') for c in df.columns]
print('Columns:', list(df.columns))

df['_NAMEKEY'] = (df['NOMBRE SEDE'] + '|' + df['CODIGO DANE SEDE'] + '|' +
                  df['NOMBRES ESTUDIANTE'] + '|' + df['GRADO'].astype(str) + '|' + df['CURSO'].astype(str))

curso_cols_g = ['NOMBRE SEDE', 'CODIGO DANE SEDE', 'CURSO', 'GRADO']

id_groups = df.groupby(['_KEY', 'NOMBRE SEDE', 'CODIGO DANE SEDE', 'GRADO', 'CURSO']).agg(
    PRUEBA=('PRUEBA', 'unique'),
    NOMBRES_EST=('NOMBRES ESTUDIANTE', 'first'),
    NAMEKEY=('_NAMEKEY', 'first'),
).reset_index()

name_groups = df.groupby(['_NAMEKEY']).agg(
    PRUEBA=('PRUEBA', 'unique'),
    NOMBRES_EST=('NOMBRES ESTUDIANTE', 'first'),
    NOMBRE_SEDE=('NOMBRE SEDE', 'first'),
    CODIGO_DANE_SEDE=('CODIGO DANE SEDE', 'first'),
    GRADO=('GRADO', 'first'),
    CURSO=('CURSO', 'first'),
    KEY_LIST=('_KEY', lambda x: list(set(x))),
).reset_index()

print(f'Total name_groups: {len(name_groups)}')
jn = name_groups[name_groups['NOMBRES_EST'].str.contains('JHANNER', na=False)]
if not jn.empty:
    r = jn.iloc[0]
    print(f'JHANNER: PRUEBA={r["PRUEBA"]}')
    completo = set(r["PRUEBA"]) >= {"MATEMATICAS", "LENGUAJE"}
    print(f'  completo={completo}')
