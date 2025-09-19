import pandas as pd

# 📥 Cargar el archivo Excel
df = pd.read_excel('DEFUNCIONES.xlsx')

# 🧼 Limpiar columna EDAD
df['EDAD'] = pd.to_numeric(df['EDAD'], errors='coerce')

# 🗓️ Convertir FECHA a datetime
df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')

# 📊 Casos por año
casos_por_año = df.groupby('AÑO')['CANT'].sum()

# 📊 Casos por mes
casos_por_mes = df.groupby('MES')['CANT'].sum()

# 📊 Ranking de delitos
ranking_delitos = df['DELITO'].value_counts()

# 📊 Promedio de edad por delito
promedio_edad_delito = df.groupby('DELITO')['EDAD'].mean()

# 📊 Casos por género
casos_por_genero = df['GENERO'].value_counts()

# 📊 Casos por rango de edad
casos_por_rango = df['RANGO DE EDAD'].value_counts()

# 📊 Casos por método
casos_por_metodo = df['METODO'].value_counts()

# 📊 Casos por vehículo
casos_por_vehiculo = df['VEHICULO'].value_counts()

# 📊 Casos por región
casos_por_region = df['REGION'].value_counts()

# 📊 Casos por comunidad
casos_por_comunidad = df['COMUNIDAD'].value_counts()

# 📤 Exportar todo a Excel con hojas separadas
with pd.ExcelWriter('informe_detallado.xlsx') as writer:
    casos_por_año.to_excel(writer, sheet_name='Casos por Año')
    casos_por_mes.to_excel(writer, sheet_name='Casos por Mes')
    ranking_delitos.to_excel(writer, sheet_name='Ranking Delitos')
    promedio_edad_delito.to_excel(writer, sheet_name='Edad Promedio por Delito')
    casos_por_genero.to_excel(writer, sheet_name='Casos por Género')
    casos_por_rango.to_excel(writer, sheet_name='Casos por Rango Edad')
    casos_por_metodo.to_excel(writer, sheet_name='Casos por Método')
    casos_por_vehiculo.to_excel(writer, sheet_name='Casos por Vehículo')
    casos_por_region.to_excel(writer, sheet_name='Casos por Región')
    casos_por_comunidad.to_excel(writer, sheet_name='Casos por Comunidad')
    df.to_excel(writer, sheet_name='Datos Originales', index=False)

print("✅ Informe completo generado: 'informe_detallado.xlsx'")
