# Automatización de Archivos de Salida NS 2020-2024

## Descripción
- Script en Python que lee `data/input/Base2.xlsx`, filtra pedidos excluidos y genera archivos Excel por proveedor con hojas `Consolidado` y `Detalle`.
- Crea además un archivo global `data/output/Resumen Proveedores NS 2020-2024.xlsx` agrupado por `Proveedor`, `Nombre` y `Año` (a partir de `Fecha Pedido`).
- Aplica formato consistente (logo, títulos, encabezados, anchos de columna, formatos numéricos y de fecha) y calcula el total de **Penalizacion con Imptos** en la hoja `Consolidado`.

## Requisitos e instalación
- Python 3.10+.
- Dependencias: `pandas`, `openpyxl`, `python-docx`.
- Instalación rápida: `pip install pandas openpyxl python-docx`.

## Entradas
- `data/input/Base2.xlsx`: fuente principal de datos.
- `data/input/Soriana-Logo.png`: logo insertado en las hojas (si no está, se avisa pero continúa).
- `EXCLUDED_PEDIDOS` en `Main.py`: lista de pedidos a excluir antes de exportar.

## Ejecución
1) Ubícate en la raíz del proyecto.
2) Ejecuta: `python Main.py`.
3) Los archivos se escriben en `data/output/`.

## Salidas generadas
- **Por proveedor**: `data/output/<proveedor nombre>/<proveedor nombre> - NS 2020-2024.xlsx` con hojas:
  - `Consolidado`: agrupación y sumatoria de montos, con un total visible arriba de la columna **Penalizacion con Imptos**.
  - `Detalle`: filas originales filtradas por proveedor/nombre.
- **Resumen global**: `data/output/Resumen Proveedores NS 2020-2024.xlsx` con la hoja `Resumen` (agrupada por `Proveedor`, `Nombre`, `Año` y sumatoria de `Penalizacion`, `IEPS`, `IVA`, `Penalizacion con Imptos`).

## Formato aplicado
- Logo y títulos (“Tiendas Soriana, S.A. de C.V.”, proveedor/nombre, “Nivel de Servicio 2020-2024”) al inicio de cada hoja.
- Encabezados con fondo verde y texto en negrilla centrado; líneas de cuadrícula ocultas.
- Formato `#,##0.00` para columnas monetarias; `dd/mm/yyyy` para `Fecha Pedido`.
- Ancho de columnas uniforme; filas de encabezado congeladas.

## Personalización rápida
- Cambiar archivo de entrada: variable `DEFAULT_INPUT_FILE` en `Main.py`.
- Ajustar exclusiones: `EXCLUDED_PEDIDOS` en `Main.py`.
- Cambiar logo: `DEFAULT_LOGO` (`data/input/Soriana-Logo.png`).

## Documentación adicional
- Se genera un Word con este resumen en `data/Documentacion_Proyecto.docx`.
