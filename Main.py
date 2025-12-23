import pandas as pd
from pathlib import Path
import unicodedata
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

DATA_DIR = Path("data")
INPUT_DIR = DATA_DIR / "input"
OUTPUT_DIR = DATA_DIR / "output"
DEFAULT_LOGO = INPUT_DIR / "Soriana-Logo.png"
DEFAULT_INPUT_FILE = INPUT_DIR / "Base1.xlsx"
SUMMARY_INPUT_FILE = INPUT_DIR / "Base1.xlsx"
SUMMARY_OUTPUT_FILE = OUTPUT_DIR / "Resumen Proveedores NS 2020-2024.xlsx"

EXCLUDED_PEDIDOS = {
    "110607909",
    "110607910",
    "121593940",
    "121647126",
    "107650539",
    "107650540",
    "108365608",
    "124886120",
    "124970463",
    "124975616",
    "124975625",
    "125778495",
    "128215996",
    "128222419",
    "128385242",
    "120579872",
    "118212854",
    "120281210",
    "120387414",
    "120397393",
    "120399826",
    "120401341",
    "120401342",
    "124505080",
    "124637548",
    "124637549",
    "124645961",
}


def sanitize_filename(text):
    """Replace characters not allowed in file/folder names."""
    invalid = '<>:"/\\|?*'
    return "".join("_" if ch in invalid else ch for ch in str(text)).strip()


def load_data(file_path, sheet_name=0):
    """Load data from an Excel file into a pandas DataFrame."""
    file_path = Path(file_path)
    try:
        print(f"Loading data from {file_path}...")
        data = pd.read_excel(file_path, sheet_name=sheet_name)
        print("Data loaded successfully.")
        return data
    except Exception as e:
        print(f"An error occurred while loading the data: {e}")
        return None


def filter_pedidos(df, excluded_pedidos, column_name="Pedido"):
    """Remove rows whose pedido value is in the excluded list."""
    if column_name not in df.columns:
        print(f"[ERROR] Column '{column_name}' not found. Available columns: {list(df.columns)}")
        return df

    before = len(df)
    filtered_df = df[~df[column_name].astype(str).isin(excluded_pedidos)]
    removed = before - len(filtered_df)
    print(f"[INFO] Removed {removed} rows with excluded pedidos. Remaining rows: {len(filtered_df)}")
    return filtered_df


def _normalize_col_name(name) -> str:
    text = str(name) if name is not None else ""
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower()
    return "".join(ch for ch in text if ch.isalnum())


def _resolve_column(df: pd.DataFrame, requested: str):
    if requested in df.columns:
        return requested

    requested_lower = str(requested).lower()
    for col in df.columns:
        if str(col).lower() == requested_lower:
            return col

    requested_norm = _normalize_col_name(requested)
    for col in df.columns:
        if _normalize_col_name(col) == requested_norm:
            return col

    return None


def _resolve_columns(df: pd.DataFrame, requested_cols):
    resolved = []
    for name in requested_cols:
        actual = _resolve_column(df, name)
        if actual is None:
            print(f"[ERROR] Column '{name}' not found. Available columns: {list(df.columns)}")
            return None
        resolved.append(actual)
    return resolved


def _to_number_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return series.fillna(0)

    s = series.astype(str).str.strip()
    s = s.replace({"-": "0", "—": "0", "": "0", "None": "0", "nan": "0"})
    s = s.str.replace("$", "", regex=False).str.replace(",", "", regex=False)
    s = s.str.replace("(", "-", regex=False).str.replace(")", "", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0)


def build_consolidado(
    df: pd.DataFrame,
    group_cols_requested,
    sum_cols_requested,
) -> pd.DataFrame | None:
    group_cols = _resolve_columns(df, group_cols_requested)
    sum_cols = _resolve_columns(df, sum_cols_requested)
    if group_cols is None or sum_cols is None:
        return None

    tmp = df.copy()
    for col in sum_cols:
        tmp[col] = _to_number_series(tmp[col])

    return tmp.groupby(group_cols, dropna=False)[sum_cols].sum().reset_index()


def build_resumen_proveedores(df: pd.DataFrame) -> pd.DataFrame | None:
    """Agrupa por proveedor, nombre y año (desde Fecha Pedido) y suma montos clave."""
    prov_col = _resolve_column(df, "Proveedor")
    nombre_col = _resolve_column(df, "Nombre")
    fecha_col = _resolve_column(df, "Fecha Pedido")
    sum_cols = _resolve_columns(
        df,
        [
            "Penalizacion",
            "IEPS",
            "IVA",
            "Penalizacion con Imptos",
        ],
    )

    if None in (prov_col, nombre_col, fecha_col) or sum_cols is None:
        return None

    tmp = df.copy()
    tmp["__year"] = pd.to_datetime(tmp[fecha_col], errors="coerce").dt.year

    for col in sum_cols:
        tmp[col] = _to_number_series(tmp[col])

    grouped = (
        tmp.groupby([prov_col, nombre_col, "__year"], dropna=False)[sum_cols]
        .sum()
        .reset_index()
    )
    grouped = grouped.rename(columns={"__year": "Año"})
    ordered_cols = [prov_col, nombre_col, "Año"] + sum_cols
    return grouped[ordered_cols]


def export_by_proveedor(df, prov_col="Proveedor", nombre_col="Nombre", output_dir=OUTPUT_DIR):
    """Split DataFrame by proveedor/nombre and save each subset to its own Excel file."""
    missing = [col for col in (prov_col, nombre_col) if col not in df.columns]
    if missing:
        print(f"[ERROR] Missing columns: {missing}. Available columns: {list(df.columns)}")
        return

    out_root = Path(output_dir)
    out_root.mkdir(parents=True, exist_ok=True)

    count = 0
    for (prov, nombre), group in df.groupby([prov_col, nombre_col], dropna=False):
        prov_str = "SIN_PROVEEDOR" if pd.isna(prov) else str(prov)
        nombre_str = "SIN_NOMBRE" if pd.isna(nombre) else str(nombre)

        folder_name = sanitize_filename(f"{prov_str} {nombre_str}")
        folder_path = out_root / folder_name
        folder_path.mkdir(parents=True, exist_ok=True)

        file_name = sanitize_filename(f"{prov_str} {nombre_str} - NS 2020-2024") + ".xlsx"
        file_path = folder_path / file_name

        consolidado = build_consolidado(
            group,
            group_cols_requested=[
                "Proveedor",
                "Nombre",
                "Division",
                "Nombre Division",
                "Tienda",
                "Pedido",
                "Fecha Pedido",
                "Organizacion",
            ],
            sum_cols_requested=[
                "Penalizacion",
                "IEPS",
                "IVA",
                "Penalizacion con Imptos",
            ],
        )

        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            if consolidado is not None:
                consolidado.to_excel(writer, index=False, sheet_name="Consolidado")
            group.to_excel(writer, index=False, sheet_name="Detalle")

        try:
            format_workbook(
                file_path,
                proveedor_display=prov_str,
                nombre_display=nombre_str,
                logo_path=DEFAULT_LOGO,
            )
        except Exception as exc:
            print(f"[WARN] Could not format {file_path}: {exc}")

        count += 1
        print(f"[INFO] Saved {len(group)} rows to {file_path}")

    print(f"[INFO] Generated {count} files in {out_root}")


def _style_headers(cells):
    """Apply Soriana header style to a row of cells."""
    fill = PatternFill(fill_type="solid", fgColor="00FD28")
    font = Font(color="000000", bold=True)
    align = Alignment(horizontal="center", vertical="center")
    for cell in cells:
        cell.fill = fill
        cell.font = font
        cell.alignment = align


def format_workbook(path, proveedor_display, nombre_display, logo_path=DEFAULT_LOGO):
    """Insert logo, titles and header styling into an Excel workbook."""
    path = Path(path)
    if not path.exists():
        print(f"[WARN] File not found for formatting: {path}")
        return

    wb = load_workbook(path)
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False

        DEFAULT_COL_WIDTH = 12  # keep columns uniform
        spacer_rows = 6  # space for logo and titles
        ws.insert_rows(1, spacer_rows)

        max_col = ws.max_column

        # Title block
        # Keep titles compact: merge only a limited span (columns 3 to 10 by default).
        merge_start = 3
        merge_end = min(max_col, 10)

        def _place_title(row, text, size=12, bold=True):
            cell = ws.cell(row=row, column=merge_start, value=text)
            cell.font = Font(size=size, bold=bold)
            cell.alignment = Alignment(horizontal="center")
            ws.merge_cells(start_row=row, start_column=merge_start, end_row=row, end_column=merge_end)

        _place_title(2, "Tiendas Soriana, S.A. de C.V.", size=14, bold=True)
        _place_title(3, f"{proveedor_display} {nombre_display}", size=13, bold=True)
        _place_title(4, "Nivel de Servicio 2020-2024", size=12, bold=True)

        # Logo
        logo_file = Path(logo_path)
        if logo_file.exists():
            try:
                img = Image(str(logo_file))
                img.width = 160
                img.height = 40
                ws.add_image(img, "B3")
            except Exception as exc:
                print(f"[WARN] Could not insert logo on {path}: {exc}")
        else:
            print(f"[WARN] Logo not found at {logo_file}")

        header_row = spacer_rows + 1
        ws.row_dimensions[header_row].height = 20
        _style_headers(ws[header_row])

        # Freeze header
        ws.freeze_panes = f"A{header_row + 1}"

        # Set a uniform column width so headers don't stretch unevenly
        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = DEFAULT_COL_WIDTH

        # Apply number/date formats based on sheet-specific columns
        header_map = {_normalize_col_name(cell.value): cell.column for cell in ws[header_row] if cell.value}

        def _apply_format(col_names, number_format):
            for name in col_names:
                col_idx = header_map.get(_normalize_col_name(name))
                if not col_idx:
                    continue
                for row in ws.iter_rows(
                    min_row=header_row + 1,
                    max_row=ws.max_row,
                    min_col=col_idx,
                    max_col=col_idx,
                ):
                    row[0].number_format = number_format

        accounting_format = "#,##0.00"
        date_format = "dd/mm/yyyy"

        if ws.title.lower() == "consolidado":
            # Place total for "Penalizacion con Imptos" near the top-right
            total_col_idx = header_map.get(_normalize_col_name("Penalizacion con Imptos"))
            if total_col_idx:
                total_col_letter = get_column_letter(total_col_idx)
                total_label_cell = ws.cell(row=3, column=total_col_idx, value="Total Penalización con Imptos")
                total_label_cell.font = Font(bold=True)
                total_label_cell.alignment = Alignment(horizontal="right")

                total_value_cell = ws.cell(row=4, column=total_col_idx)
                total_value_cell.value = f"=SUM({total_col_letter}{header_row + 1}:{total_col_letter}{ws.max_row})"
                total_value_cell.number_format = accounting_format
                total_value_cell.alignment = Alignment(horizontal="right")

            _apply_format(
                [
                    "Penalizacion",
                    "IEPS",
                    "IVA",
                    "Penalizacion con Imptos",
                ],
                accounting_format,
            )
            _apply_format(["Fecha Pedido"], date_format)
        elif ws.title.lower() == "detalle":
            _apply_format(
                [
                    "fact_surt",
                    "Costo",
                    "Precio Vta",
                    "Mgen",
                    "Penalizacion",
                    "IEPS",
                    "IVA",
                    "Penalizacion con Imptos",
                ],
                accounting_format,
            )
            _apply_format(["Fecha Pedido"], date_format)
        elif ws.title.lower() == "resumen":
            _apply_format(
                [
                    "Penalizacion",
                    "IEPS",
                    "IVA",
                    "Penalizacion con Imptos",
                ],
                accounting_format,
            )

    wb.save(path)


def generate_resumen_proveedores(
    input_file=SUMMARY_INPUT_FILE,
    output_file=SUMMARY_OUTPUT_FILE,
    logo_path=DEFAULT_LOGO,
    excluded_pedidos=EXCLUDED_PEDIDOS,
):
    """Crear el archivo 'Resumen Proveedores NS 2020-2024' directamente desde Base2.xlsx."""
    data = load_data(input_file)
    if data is None:
        return

    if excluded_pedidos:
        data = filter_pedidos(data, excluded_pedidos, column_name="Pedido")

    resumen = build_resumen_proveedores(data)
    if resumen is None:
        print("[ERROR] No se pudo construir el resumen de proveedores (faltan columnas).")
        return

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        resumen.to_excel(writer, index=False, sheet_name="Resumen")

    try:
        format_workbook(
            output_file,
            proveedor_display="Resumen Proveedores",
            nombre_display="NS 2020-2024",
            logo_path=logo_path,
        )
    except Exception as exc:
        print(f"[WARN] No se pudo formatear el resumen: {exc}")

    print(f"[INFO] Resumen guardado en {output_file}")


if __name__ == "__main__":
    file_path = DEFAULT_INPUT_FILE
    data = load_data(file_path)
    if data is not None:
        filtered_data = filter_pedidos(data, EXCLUDED_PEDIDOS, column_name="Pedido")
        print(filtered_data.head())
        export_by_proveedor(filtered_data, prov_col="Proveedor", nombre_col="Nombre", output_dir=OUTPUT_DIR)

    generate_resumen_proveedores()
