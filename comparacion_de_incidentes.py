import pandas as pd

# 1. Cargar archivos
df_inc = pd.read_csv('Incidentes_registrados.csv')
df_cmca = pd.read_excel('Copia de CMCA_2026.xlsx', sheet_name='Sheet1', header=2)

# Convertir fechas del Excel y quitar filas vacías (como la fila de unidades)
df_cmca['DateTime'] = pd.to_datetime(df_cmca['Date & Time'], format='%d/%m/%Y %H:%M', errors='coerce')
df_cmca = df_cmca[df_cmca['DateTime'].notna()].copy()

# Convertir fechas del CSV redondeando a la hora en punto (.dt.floor('h')) para que 14:50 coincida con 14:00
df_inc['Inicio'] = pd.to_datetime(df_inc['Fecha inicio'], format='%d/%m/%Y %H:%M', errors='coerce').dt.floor('h')
df_inc['Fin'] = pd.to_datetime(df_inc['Fecha fin'], format='%d/%m/%Y %H:%M', errors='coerce').dt.floor('h')

# Si la fecha fin dice "Por definir" (NaT), abarca hasta el final del Excel si hay equipos afectados
max_fecha = df_cmca['DateTime'].max()
df_inc['Fin'] = df_inc.apply(
    lambda r: max_fecha if pd.isna(r['Fin']) and pd.notna(r['Equipos afectados (normalizados)']) else (r['Inicio'] if pd.isna(r['Fin']) else r['Fin']),
    axis=1
)

# 2. Relacionar cada parámetro con su columna Status
param_status = {}
for col in df_cmca.columns:
    if 'Status' in str(col):
        idx = df_cmca.columns.get_loc(col)
        param = str(df_cmca.columns[idx - 1]).strip()
        param_status[param] = col
        # Si la celda está vacía o tiene '---' (ej. corte eléctrico), poner 'NoData'
        df_cmca[col] = df_cmca[col].fillna('NoData').astype(str).str.strip().replace({'---': 'NoData', 'nan': 'NoData'})

# Función para saber si una fecha y parámetro están dentro de algún incidente
def esta_registrado(fecha, param):
    for _, inc in df_inc.iterrows():
        if pd.isna(inc['Equipos afectados (normalizados)']):
            continue
        equipos = [e.strip() for e in str(inc['Equipos afectados (normalizados)']).split(',')]
        if param in equipos and (inc['Inicio'] <= fecha <= inc['Fin']):
            return True
    return False

# -------------------------------------------------------------
# 1. Banderas distintas de 'Ok' que NO están registradas
# -------------------------------------------------------------
no_registrados = []
for param, col in param_status.items():
    sub = df_cmca[df_cmca[col] != 'Ok'][['DateTime', col]]
    for _, row in sub.iterrows():
        if not esta_registrado(row['DateTime'], param):
            no_registrados.append({
                'Parámetro': param,
                'Bandera': row[col],
                'Fecha': row['DateTime']
            })

df_no_reg = pd.DataFrame(no_registrados)

print("=== 1. BANDERAS (APARTE DE OK) NO REGISTRADAS ===")
if not df_no_reg.empty:
    df_no_reg = df_no_reg.sort_values(['Parámetro', 'Bandera', 'Fecha']).reset_index(drop=True)
    # Agrupar horas seguidas del mismo parámetro y bandera
    cambio = (
        (df_no_reg['Parámetro'] != df_no_reg['Parámetro'].shift(1)) |
        (df_no_reg['Bandera'] != df_no_reg['Bandera'].shift(1)) |
        (df_no_reg['Fecha'] != df_no_reg['Fecha'].shift(1) + pd.Timedelta(hours=1))
    )
    df_no_reg['Grupo'] = cambio.cumsum()
    
    resumen = df_no_reg.groupby(['Parámetro', 'Bandera', 'Grupo'], as_index=False).agg(
        Inicio=('Fecha', lambda x: x.min().strftime('%d/%m/%Y %H:%M')),
        Fin=('Fecha', lambda x: x.max().strftime('%d/%m/%Y %H:%M')),
        Horas=('Fecha', 'count')
    ).drop(columns=['Grupo'])
    
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 150)
    print(resumen.to_string(index=False))
else:
    print("No hay banderas sin registrar.")

# -------------------------------------------------------------
# 2. Comparación de los incidentes con el Excel en ese lapso
# -------------------------------------------------------------
print("\n=== 2. COMPARACIÓN DE INCIDENTES CON BANDERAS DEL 2026 EN ESE LAPSO ===")
comparacion = []

for _, inc in df_inc.iterrows():
    sub_excel = df_cmca[(df_cmca['DateTime'] >= inc['Inicio']) & (df_cmca['DateTime'] <= inc['Fin'])]
    equipos = [e.strip() for e in str(inc['Equipos afectados (normalizados)']).split(',')] if pd.notna(inc['Equipos afectados (normalizados)']) else []
    
    banderas_lapso = set()
    for eq in equipos:
        if eq in param_status:
            col = param_status[eq]
            banderas_lapso.update(sub_excel[col].unique())
            
    comparacion.append({
        'N°': inc['N°'],
        'Inicio': inc['Fecha inicio'],
        'Fin': inc['Fecha fin'],
        'Parámetros': inc['Equipos afectados (normalizados)'] if pd.notna(inc['Equipos afectados (normalizados)']) else 'Ninguno',
        'Bandera CSV': inc['Bandera del SAD'] if pd.notna(inc['Bandera del SAD']) else 'Sin bandera',
        'Banderas en Excel': ', '.join(sorted(banderas_lapso)) if banderas_lapso else 'Sin equipos'
    })

df_comp = pd.DataFrame(comparacion)
print(df_comp.to_string(index=False))