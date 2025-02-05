import pandas as pd
import matplotlib.pyplot as plt

# Leer archivo
df = pd.read_excel("data.xlsx", engine="openpyxl")

# Limpiar datos y convertir columnas numéricas
df = df.apply(lambda x: x.astype(str).str.strip() if x.dtype == "object" else x)
df[['VENTA_C1', 'VENTA_C2', 'COSTO', 'CANTIDAD']] = df[['VENTA_C1', 'VENTA_C2', 'COSTO', 'CANTIDAD']].apply(pd.to_numeric, errors='coerce').fillna(0)

# Calcular métricas antes del filtrado
df['Ventas Totales'] = df['VENTA_C1'] + df['VENTA_C2']
df['Margen'] = round(df['Ventas Totales'] - df['COSTO'], 2)
df['Rentabilidad'] = df.apply(lambda row: round((row['Margen'] / row['Ventas Totales']) * 100, 2) if row['Ventas Totales'] != 0 else 0, axis=1)

# Filtrar DataFrames según las condiciones
df_ambito = df[(df['UNIDAD DE NEGOCIO'] == 'RETAIL') & (df['ZONA'] != 'MDP CENTRO')]
df_centro = df[df['UNIDAD DE NEGOCIO'] == 'ECOMMERCE']
df_vep = df[df['UNIDAD DE NEGOCIO'] == 'VEP']
df_ecommerce = df[(df['UNIDAD DE NEGOCIO'] == 'RETAIL') & (df['ZONA'] == 'MDP CENTRO')]

def calcular_metricas(data_frame):
    resultados = []

    # Iterar por rubro y subrubro para organizar las métricas por columnas
    for rubro in data_frame['RUBRO'].unique():
        for subrubro in data_frame[data_frame['RUBRO'] == rubro]['SUBRUBRO'].unique():
            df_rubro_subrubro = data_frame[(data_frame['RUBRO'] == rubro) & (data_frame['SUBRUBRO'] == subrubro)]

            # Métricas
            sumatoria_articulos = df_rubro_subrubro['CANTIDAD'].sum()
            articulo_mas_vendido = df_rubro_subrubro['DETALLE'].mode().iloc[0] if not df_rubro_subrubro['DETALLE'].mode().empty else 'Sin datos'
            ganancia_por_subrubro = df_rubro_subrubro['Ventas Totales'].sum()
            margen_total = df_rubro_subrubro['Margen'].sum()
            rentabilidad_promedio = df_rubro_subrubro['Rentabilidad'].mean()

            # Agregar las métricas como una fila organizada en columnas
            resultados.append([rubro, subrubro, sumatoria_articulos, articulo_mas_vendido, ganancia_por_subrubro, margen_total, round(rentabilidad_promedio, 2)])

    columnas = ['Rubro', 'Subrubro', 'Cantidad Artículos Vendidos', 'Artículo Más Vendido', 'Ganancia por Subrubro', 'Margen Total', 'Rentabilidad Promedio (%)']
    return pd.DataFrame(resultados, columns=columnas)

# Calcular las métricas para cada DataFrame
df_ambito_metrics = calcular_metricas(df_ambito)
df_centro_metrics = calcular_metricas(df_centro)
df_vep_metrics = calcular_metricas(df_vep)
df_ecommerce_metrics = calcular_metricas(df_ecommerce)

def crear_grafico_torta(df_metrics, nombre_grafico):
    if df_metrics.empty:
        return

    # Iterar por cada rubro y crear un gráfico
    for rubro in df_metrics['Rubro'].unique():
        df_rubro = df_metrics[df_metrics['Rubro'] == rubro]

        # Filtrar valores negativos
        df_rubro = df_rubro[df_rubro['Ganancia por Subrubro'] > 0]

        if df_rubro.empty or df_rubro['Ganancia por Subrubro'].sum() == 0:
            continue

        plt.figure(figsize=(8, 8))
        plt.pie(df_rubro['Ganancia por Subrubro'], labels=df_rubro['Subrubro'], autopct='%1.1f%%', startangle=140)
        plt.title(f"Distribución de Ganancias por Subrubro ({nombre_grafico} - {rubro})")
        plt.savefig(f"{nombre_grafico}_{rubro}_Grafico_Torta.png")
        plt.close()

# Crear gráficos para cada conjunto de datos
crear_grafico_torta(df_ambito_metrics, "Ambito")
crear_grafico_torta(df_centro_metrics, "Centro")
crear_grafico_torta(df_vep_metrics, "VEP")
crear_grafico_torta(df_ecommerce_metrics, "Ecommerce")

# Exportar a Excel solo si los DataFrames no están vacíos
with pd.ExcelWriter("Resultados_Métricas.xlsx") as writer:
    if not df_ambito_metrics.empty:
        df_ambito_metrics.to_excel(writer, sheet_name="Ámbito", index=False)
    if not df_centro_metrics.empty:
        df_centro_metrics.to_excel(writer, sheet_name="Centro", index=False)
    if not df_vep_metrics.empty:
        df_vep_metrics.to_excel(writer, sheet_name="VEP", index=False)
    if not df_ecommerce_metrics.empty:
        df_ecommerce_metrics.to_excel(writer, sheet_name="Ecommerce", index=False)

print("Archivo Excel generado exitosamente.")