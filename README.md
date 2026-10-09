# Proyecto Análisis Excel: dashboard automático

Convierte **cualquier** archivo de datos (Excel o CSV) en un dashboard de Excel con
gráficas de análisis modernas, sin configurar nada.

```bash
pip install -r requirements.txt
python dashboard_excel.py mis_datos.xlsx        # crea mis_datos_dashboard.xlsx
```

En Windows también puedes **arrastrar el archivo sobre `crear_dashboard.bat`**, o ejecutar
`python dashboard_excel.py` sin argumentos para que se abra un selector de archivo.

## Qué genera

| Hoja | Contenido |
|---|---|
| **Dashboard** | Tarjetas KPI y paneles de gráficas listos para presentar o imprimir |
| **Datos** | Tus datos limpios como Tabla de Excel |
| **Calculos** | Agregaciones con fórmulas (`SUMIFS`/`COUNTIFS`) que alimentan las gráficas |
| **Acerca de** | Qué columnas se detectaron, con qué rol y para qué se usaron |

Visualizaciones del Dashboard (se incluyen solo las que aplican a tus datos):

- **Tarjetas KPI**: total, promedio, registros y variación del último período completo (▲/▼ con color).
- **Tendencia** (área): agrupa sola por día, semana, mes, trimestre o año según el rango de fechas.
- **Ranking** (barras horizontales ordenadas) con etiquetas de valor y el líder en el subtítulo.
- **Participación** (dona): top 5 + "Otros".
- **Composición en el tiempo** (columnas apiladas por categoría).
- **Distribución** (histograma con rangos redondeados).
- **Mapa de calor**: categoría × categoría, o año × mes (estacionalidad) si no hay segunda categoría.
- **Dispersión** entre las dos medidas numéricas.
- **Tabla de detalle** con barras de datos.

Los colores usan una paleta validada para daltonismo, con un orden fijo. Las barras y las áreas
siempre arrancan en cero y el gris se reserva para "Otros".

## Cómo funciona "sin importar el Excel"

El script detecta solo:

- la **hoja con más datos** (o la que indiques con `--hoja`);
- la **fila de encabezados**, aunque haya títulos o filas vacías arriba;
- **filas de totales** ("Total", "Total general") y las descarta para no contarlas doble;
- **números en cualquier formato**: `1.234,56`, `$1,234.56`, `(500)`, `45%`, `COP 2.000`;
- **fechas** como fecha real o como texto (`dd/mm/aaaa`, ISO…);
- columnas de **identificadores** (ID, código, NIT, SKU, factura…), que se ignoran;
- columnas de **año** (2023, 2024…), que se tratan como categoría y no se suman;
- la **medida principal** (prefiere ventas/ingresos/monto/total y evita sumar precios
  unitarios, porcentajes o calificaciones) y las **categorías** más útiles.

Si no hay ninguna columna numérica, el dashboard cuenta registros.

## Cambiar lo que eligió

```bash
python dashboard_excel.py datos.xlsx --medida "Ventas" --categoria "Región" --fecha "Fecha pedido"
python dashboard_excel.py datos.xlsx --hoja "Hoja2" --titulo "Ventas Q3" -o reporte.xlsx
```

## Actualizar los datos

Todo el Dashboard son fórmulas sobre la hoja **Datos**. Si corriges valores en esa hoja, las gráficas
se recalculan solas. Si llegan datos nuevos con más filas, vuelve a ejecutar el script sobre el archivo nuevo.

## Ejemplos y pruebas

```bash
python ejemplos/generar_ejemplos.py     # crea ventas.xlsx, encuesta.csv, inventario.xlsx
python dashboard_excel.py ejemplos/ventas.xlsx
python -m pytest -q tests
```

`ejemplos/ventas_dashboard.xlsx` es un resultado de muestra que puedes abrir directamente.

## Limitaciones

- La dispersión usa como máximo los primeros 5.000 registros.
- Los archivos `.xls` antiguos requieren `xlrd` (incluido en `requirements.txt`).
- Excel no permite crear segmentaciones (slicers) desde Python. Si las necesitas, agrégalas
  sobre la Tabla de la hoja **Datos**.
