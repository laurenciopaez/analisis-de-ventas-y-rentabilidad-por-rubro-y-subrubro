# Análisis de Ventas y Rentabilidad por Rubro y Subrubro

Este proyecto realiza un análisis detallado de las ventas, márgenes y rentabilidad de diferentes unidades de negocio en función de rubros y subrubros. A partir de un archivo de datos en formato Excel (`data.xlsx`), se calculan diversas métricas clave y se generan gráficos de torta para representar la distribución de ganancias por subrubro. Los resultados se exportan a un archivo Excel denominado `Resultados_Métricas.xlsx`.

Funcionalidades
El programa realiza las siguientes acciones:

Leer el archivo Excel: Se carga el archivo data.xlsx y se limpia para convertir las columnas numéricas en el formato adecuado.
Calcular métricas clave:
Ventas Totales: Suma de las ventas de dos columnas.
Margen: Diferencia entre las ventas totales y el costo.
Rentabilidad: Rentabilidad calculada como margen dividido entre ventas totales, expresada como porcentaje.
Filtrar por unidades de negocio: Los datos se filtran en 4 unidades de negocio específicas: RETAIL, ECOMMERCE, VEP, y MDP CENTRO.
Calcular métricas por Rubro y Subrubro: Se calculan diversas métricas como:
Cantidad de artículos vendidos.
Artículo más vendido.
Ganancia por subrubro.
Margen total.
Rentabilidad promedio.
Generar gráficos de torta: Para cada conjunto de métricas calculadas, se genera un gráfico de torta representando la distribución de las ganancias por subrubro.
Exportar resultados: Los resultados se exportan a un archivo Excel denominado Resultados_Métricas.xlsx.
Uso
Preparación del archivo de entrada:

Asegúrate de tener un archivo data.xlsx con las siguientes columnas:
VENTA_C1, VENTA_C2, COSTO, CANTIDAD, UNIDAD DE NEGOCIO, ZONA, RUBRO, SUBRUBRO, DETALLE.
Ejecuta el script:

Corre el script Python que realiza el análisis.
Verifica los resultados:

Los resultados de las métricas calculadas serán exportados a un archivo Excel.
Los gráficos de torta generados se guardarán como imágenes PNG en el directorio de trabajo.

## Requisitos

Para ejecutar este proyecto, necesitas tener instalados los siguientes paquetes:

- Python 3.x
- Pandas
- Matplotlib
- openpyxl

Puedes instalarlos utilizando pip:

```bash
pip install pandas matplotlib openpyxl
