# Análisis de Ventas y Rentabilidad por Rubro y Subrubro

Este proyecto realiza un análisis detallado de las ventas, márgenes y rentabilidad de diferentes unidades de negocio en función de rubros y subrubros. A partir de un archivo de datos en formato Excel (`data.xlsx`), se calculan métricas clave como ventas totales, margen y rentabilidad, y se generan gráficos de torta para visualizar la distribución de ganancias por subrubro. Los resultados obtenidos se exportan a un archivo Excel llamado `Resultados_Métricas.xlsx`, junto con las imágenes de los gráficos generados.

## Funcionalidades

El programa lleva a cabo las siguientes acciones:

1. **Lectura y limpieza de datos**:
   - Carga el archivo `data.xlsx`.
   - Convierte las columnas numéricas en el formato adecuado.
   - Limpia los valores de las columnas de texto (eliminación de espacios extra).

2. **Cálculo de métricas clave**:
   - **Ventas Totales**: Suma de las ventas de las columnas `VENTA_C1` y `VENTA_C2`.
   - **Margen**: Diferencia entre las ventas totales y el costo (`VENTA_C1 + VENTA_C2 - COSTO`).
   - **Rentabilidad**: Calculada como `(Margen / Ventas Totales) * 100`, expresada en porcentaje.

3. **Filtrado de datos**:
   - Los datos se filtran según las unidades de negocio: **RETAIL**, **ECOMMERCE**, **VEP**, y **MDP CENTRO**.

4. **Cálculo de métricas por rubro y subrubro**:
   - **Cantidad de artículos vendidos**: Suma de la columna `CANTIDAD`.
   - **Artículo más vendido**: El artículo con mayor frecuencia en la columna `DETALLE`.
   - **Ganancia por subrubro**: La suma de las ganancias totales por cada subrubro.
   - **Margen total**: Suma de los márgenes de cada subrubro.
   - **Rentabilidad promedio**: Promedio de rentabilidad por subrubro.

5. **Generación de gráficos**:
   - Se genera un gráfico de torta que representa la distribución de las ganancias por subrubro para cada unidad de negocio.

6. **Exportación de resultados**:
   - Los resultados de las métricas calculadas se exportan a un archivo Excel llamado `Resultados_Métricas.xlsx`.
   - Los gráficos de torta generados se guardan como imágenes PNG en el directorio de trabajo.

## Uso

### Preparación del archivo de entrada:

Asegúrate de tener un archivo `data.xlsx` con las siguientes columnas:
- `VENTA_C1`
- `VENTA_C2`
- `COSTO`
- `CANTIDAD`
- `UNIDAD DE NEGOCIO`
- `ZONA`
- `RUBRO`
- `SUBRUBRO`
- `DETALLE`

### Ejecución del script:

Corre el script Python que realiza el análisis y genera los resultados.

### Verificación de resultados:

- Los resultados de las métricas calculadas serán exportados a un archivo Excel llamado `Resultados_Métricas.xlsx`.
- Los gráficos de torta generados se guardarán como imágenes PNG en el directorio de trabajo.

## Requisitos

Para ejecutar este proyecto, necesitas tener instalados los siguientes paquetes:

- Python 3.x
- Pandas
- Matplotlib
- openpyxl

Puedes instalarlos utilizando `pip`:

```bash
pip install pandas matplotlib openpyxl
