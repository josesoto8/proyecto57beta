import pandas as pd

# 1. Definición de archivos y hoja
archivo_origen = "Copia de Excel_2025_con_Limpios_FINAL.xlsx"
archivo_salida = "datosLPZ2025.xlsx"
hoja_origen = "Verificación"

# 2. Cargar los datos desde la hoja 'Verificación'
df_origen = pd.read_excel(archivo_origen, sheet_name=hoja_origen)

# 3. Diccionario de mapeo: 'columna_destino': 'columna_origen'
mapeo_columnas = {
    'date': 'Date & Time',
    'ws': 'VV_valorvalidado',
    'wd': 'DV_valorvalidado',
    'HR': 'HR_valorvalidado',
    'PB': 'PB_valorvalidado',
    'Prec': 'Precip_valorvalidado',
    'UVB': 'UVB_valorvalidado',
    'UVI': 'UVI_valorvalidado',
    'TMP': 'Temp_valorvalidado',
    'O3': 'O3_valorvalidado',
    'NO': 'NO_valorvalidado',
    'NO2': 'NO2_valorvalidado',
    'NOX': 'NOX_valorvalidado',
    'CO': 'CO_valorvalidado',
    'PM10': 'PM10_valorvalidado',
    'PM2.5': 'PM2_5_valorvalidado',
    'NO2': 'NO2_valorvalidado',
    'SO2': 'SO2_valorvalidado'
    ''
}

# 4. Crear el DataFrame final
df_ppp2025 = pd.DataFrame()

# Parsear fecha y hora al formato estandarizado
df_ppp2025['date'] = pd.to_datetime(df_origen[mapeo_columnas['date']], dayfirst=True)

# Mapear y convertir datos numéricos; reemplazar casillas en blanco/NaN con -99
for col_destino, col_origen in list(mapeo_columnas.items())[1:]:
    df_ppp2025[col_destino] = pd.to_numeric(df_origen[col_origen], errors='coerce').fillna(-99)

# 5. Exportar a Excel con la estructura de ppp.xlsx para R (openair)
df_ppp2025.to_excel(archivo_salida, index=False)
print(f"Archivo generado correctamente: {archivo_salida}")