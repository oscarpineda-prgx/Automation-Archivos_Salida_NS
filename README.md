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
- `EXCLUDED_PEDIDOS` en `src/Main.py`: lista de pedidos a excluir antes de exportar.

## Ejecución
1) Ubícate en la raíz del proyecto.
2) Ejecuta: `python src/Main.py`.
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
- Cambiar archivo de entrada: variable `DEFAULT_INPUT_FILE` en `src/Main.py`.
- Ajustar exclusiones: `EXCLUDED_PEDIDOS` en `src/Main.py`.
- Cambiar logo: `DEFAULT_LOGO` (`data/input/Soriana-Logo.png`).

## Los tres scripts

| Script | Qué hace |
|---|---|
| `src/Main.py` | Genera los archivos por proveedor del periodo **2020-2024** (lee `data/input/Base1.xlsx`). |
| `src/Main_2025.py` | Lo mismo para **enero-agosto 2025** (lee `data/input/Base 2025.xlsx`). |
| `src/Main_Dir.py` | Solo crea el árbol de carpetas por proveedor en `data/output/carpetas/`, sin generar Excel. |

## Estructura

```
src/            los tres scripts
data/
  input/        bases de origen y el logo
  output/       salidas generadas (no se versionan: son GB por corrida)
entregables/    entregas ya enviadas
  0 - NS 2024/            (1.2 GB)
  0 - NS ene-ago 2025/
docs/           Documentacion_Proyecto.docx
```

Las rutas se calculan desde la raíz del proyecto, así que los scripts se pueden
correr desde cualquier carpeta.

## Nota sobre el repositorio

Hasta este ordenamiento, el repo versionaba `data/output/` —617 MB de Excel
generados— mientras que `Main_2025.py` y `Main_Dir.py` **no estaban versionados**.
Se invirtió: ahora se versiona el código y las salidas quedan fuera con
`.gitignore`. Los archivos siguen en disco; solo dejaron de rastrearse.

Los Excel ya subidos siguen en el historial de GitHub, así que el repo remoto
conserva su tamaño. Limpiarlo de verdad exige reescribir el historial, que es
una operación aparte y hay que acordarla antes de hacerla.

## Documentación adicional
- Se genera un Word con este resumen en `docs/Documentacion_Proyecto.docx`.
