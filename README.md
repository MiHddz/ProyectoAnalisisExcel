# Proyecto Análisis Excel: dashboard automático

Hay dos formas de usarlo:

| | **Plantilla de Excel** (sin instalar nada) | **Generador en Python** |
|---|---|---|
| Archivo | `plantilla/Plantilla_Dashboard.xlsx` | `dashboard_excel.py` |
| Requiere | Solo Excel 2010 o superior (recomendado Microsoft 365) | Python + `pip install` |
| Cómo se usa | Pegas tus datos en la hoja **Datos** | `python dashboard_excel.py archivo.xlsx` |
| Ideal para | Uso diario con bases pequeñas y medianas | Archivos grandes, reportes con títulos y totales |

## Plantilla de Excel (sin instalar nada)

1. Descarga `plantilla/Plantilla_Dashboard.xlsx` desde GitHub (abre el archivo y usa el botón **Download raw file**).
2. Ábrela en Excel. Si aparece "Habilitar edición", haz clic. **No tiene macros**.
3. Ve a la hoja **Datos**, presiona **Ctrl+A** y luego **Supr** para borrar el ejemplo.
4. Pega tu base de datos desde la celda **A1** (encabezados en la fila 1, un registro por fila).
5. Vuelve a la hoja **Dashboard**: todo se recalcula solo.

La hoja **Config** muestra qué columna eligió para cada cosa (medida, fecha, categorías, etapa,
departamento, costo). Si quieres otra, escógela en la lista desplegable amarilla. Para volver a la
detección automática, borra la celda amarilla.

Visualizaciones incluidas:

| Visualización | Qué muestra |
|---|---|
| Tarjetas KPI | Total, registros, promedio y margen (o valores distintos), con ▲/▼ frente al mes anterior |
| Líneas | Tendencia; se agrupa sola por día, semana, mes, trimestre o año según el rango de fechas |
| Barras horizontales | Ranking top 10 de la categoría principal y de la secundaria |
| Dona | Participación: top 5 + "Otros" (usa la categoría con menos valores) |
| Barras apiladas | Últimos 12 meses divididos por categoría |
| Matriz (tipo tabla dinámica) | Categoría × últimos 6 meses, con color según el valor y total con barras de datos |
| Mapa | Mosaico de los 33 departamentos de Colombia (acepta nombres con o sin tildes) + top 5 |
| Mapa de calor | Hora × día de la semana (o mes × día si la fecha no tiene hora) |
| Dispersión | Relación entre las dos medidas numéricas |
| Embudo | Registros por etapa (ej. Visita → Carrito → Pago → Compra) con conversión total |

Si a tus datos les falta algo (por ejemplo, no hay fecha), ese panel muestra un aviso ⚠ que dice
qué columna elegir en **Config**, y el resto del tablero sigue funcionando.

**Bases sin dinero (envíos, tickets, encuestas, inventarios de registros):** si ninguna columna parece
ventas, monto, valor o cantidad, la plantilla **cuenta registros** en lugar de sumar. Nunca suma columnas
como `GUIA`, `ORDEN`, `CUENTA`, `CEDULA` o `TELEFONO`. Por ejemplo, en una base de mensajería las tarjetas
muestran: registros (guías), promedio de `DIAS ENTREGA`, % del estado más común (por ejemplo, `% ENTREGADO`)
y productos distintos. El embudo usa la columna `ESTADO` y el mapa usa `DEPARTAMENTO` (en mayúsculas,
sin tildes o con tildes).

**Si algo sale en blanco:**
- Archivo descargado: haz clic en **Habilitar edición** (barra amarilla). La plantilla trae los resultados
  del ejemplo guardados, pero con tus datos Excel necesita calcular.
- Fechas guardadas como texto: arriba del tablero aparece un aviso ⚠ con el nombre de la columna. Selecciónala
  y usa **Datos › Texto en columnas › Siguiente › Siguiente › Fecha: DMA › Finalizar**.
- Cálculo manual: en **Fórmulas › Opciones para el cálculo** elige **Automático** (o presiona Ctrl+Alt+F9).

Límites de la plantilla: analiza hasta 40 columnas; en cada categoría busca hasta 40 valores distintos
(10 en el ranking); la dispersión usa los primeros 5.000 registros. Si tu base tiene cientos de miles de filas,
puede tardar unos segundos en recalcular: para esos casos usa el generador en Python.

Para regenerar la plantilla después de modificar su diseño: `python plantilla/generar_plantilla.py`. Con
`--valores copia_recalculada.xlsx` se guardan los resultados calculados dentro del archivo (para que se vea completa
en Vista protegida).

## Generador en Python

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
