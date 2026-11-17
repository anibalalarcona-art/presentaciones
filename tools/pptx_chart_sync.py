#!/usr/bin/env python3
"""
Sincroniza los gráficos de un .pptx con los datos vivos de su libro Excel fuente.

Cada gráfico del .pptx (ppt/charts/chartN.xml) guarda, además de un vínculo
externo al Excel original, una "caché" de los valores que efectivamente se
dibujan (<c:numCache>/<c:strCache>). PowerPoint sólo refresca esa caché si
abres el archivo conectado a la red donde vive el Excel. Este script hace lo
mismo sin PowerPoint: lee el Excel, encuentra para cada serie el rango de
celdas que la alimenta (su fórmula <c:f>, p.ej. "Combustibles!$B$3:$B$48"),
toma los valores actuales de esas celdas (extendiendo el rango si hay filas o
columnas nuevas con datos) y reescribe sólo esa caché dentro del XML del
gráfico. Todo lo demás del .pptx (diseño, colores, fuentes, texto, capas)
queda byte a byte intacto.

Uso:
    python3 pptx_chart_sync.py --xlsx "Graficos.xlsx" --pptx "Laminas.pptx" --out "Laminas.pptx"

Sin --out, escribe "<nombre>.synced.pptx" junto al original.
"""
import argparse
import re
import shutil
import sys
import zipfile
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

EXCEL_EPOCH = date(1899, 12, 30)
MAX_SCAN = 600  # cuántas filas/columnas más allá del rango actual se exploran buscando datos nuevos

CELL_REF_RE = re.compile(
    r"^(?:'([^']+)'|([^!']+))!\$([A-Za-z]+)\$(\d+)(?::\$([A-Za-z]+)\$(\d+))?$"
)


class RefParseError(ValueError):
    pass


def to_serial(value):
    """Convierte fechas de Excel/Python al número de serie que usa el caché del gráfico."""
    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        return (value - EXCEL_EPOCH).days
    return value


def fmt_number(value):
    if isinstance(value, float) and value.is_integer() and abs(value) < 1e15:
        return str(int(value))
    return repr(value) if isinstance(value, float) else str(value)


def parse_ref(ref):
    m = CELL_REF_RE.match(ref)
    if not m:
        raise RefParseError(f"No pude interpretar la referencia de celda: {ref!r}")
    sheet = m.group(1) or m.group(2)
    c1, r1 = m.group(3), int(m.group(4))
    if m.group(5):
        c2, r2 = m.group(5), int(m.group(6))
    else:
        c2, r2 = c1, r1
    return sheet, c1, r1, c2, r2


def format_ref(sheet, c1, r1, c2, r2):
    needs_quote = not re.match(r"^[^\W\d][\w]*$", sheet, re.UNICODE)
    s = f"'{sheet}'" if needs_quote else sheet
    if c1 == c2 and r1 == r2:
        return f"{s}!${c1}${r1}"
    return f"{s}!${c1}${r1}:${c2}${r2}"


def is_row_range(c1, r1, c2, r2):
    return r1 == r2 and c1 != c2


def is_col_range(c1, r1, c2, r2):
    return c1 == c2 and r1 != r2


def cell_value(ws, col, row):
    return ws[f"{col}{row}"].value


def scan_extent(ws, sheet_name, c1, r1, c2, r2, max_scan=MAX_SCAN):
    """
    A partir del rango declarado, devuelve el nuevo extremo final.

    Dos pasos, con tolerancia distinta a propósito:
    1) RECORTAR dentro del rango ya declarado: toma el último índice no
       vacío entre el inicio y el final declarado, tolerando huecos
       internos (p.ej. un pronóstico que aún no publica una institución
       un mes dado) — así no se corta la serie por un dato faltante
       intermedio.
    2) CRECER más allá del final declarado: sólo avanza mientras las
       celdas siguientes estén seguidas, sin tolerar ningún hueco. Una
       sola celda vacía detiene el crecimiento. Esto es a propósito más
       estricto que el recorte: una hoja puede tener, en la misma fila o
       columna, OTRO bloque de datos no relacionado un poco más allá (p.ej.
       celdas de apoyo para otro cálculo) separado por una sola celda en
       blanco — tolerar huecos al crecer terminaría incorporando ese
       bloque ajeno al rango del gráfico.
    """
    if is_row_range(c1, r1, c2, r2) or (c1 == c2 and r1 == r2):
        start_idx = column_index_from_string(c1)
        end_idx = column_index_from_string(c2)
        trimmed_end = start_idx
        for i in range(start_idx, end_idx + 1):
            if cell_value(ws, get_column_letter(i), r1) is not None:
                trimmed_end = i
        grown_end = end_idx
        for i in range(end_idx + 1, end_idx + max_scan + 1):
            if cell_value(ws, get_column_letter(i), r1) is not None:
                grown_end = i
            else:
                break
        return c1, r1, get_column_letter(max(trimmed_end, grown_end)), r2
    else:
        trimmed_end = r1
        for r in range(r1, r2 + 1):
            if cell_value(ws, c1, r) is not None:
                trimmed_end = r
        grown_end = r2
        for r in range(r2 + 1, r2 + max_scan + 1):
            if cell_value(ws, c1, r) is not None:
                grown_end = r
            else:
                break
        return c1, r1, c2, max(trimmed_end, grown_end)


def read_values(ws, c1, r1, c2, r2):
    if is_row_range(c1, r1, c2, r2) or (c1 == c2 and r1 == r2):
        out = []
        for i in range(column_index_from_string(c1), column_index_from_string(c2) + 1):
            out.append(cell_value(ws, get_column_letter(i), r1))
        return out
    else:
        out = []
        for r in range(r1, r2 + 1):
            out.append(cell_value(ws, c1, r))
        return out


def build_cache_xml(tag, format_code, values, is_date=False):
    """tag es 'c:numCache' o 'c:strCache'."""
    pts = []
    for i, v in enumerate(values):
        if v is None or v == "":
            continue
        if is_date:
            v = to_serial(v)
        v_str = fmt_number(v) if not isinstance(v, str) else v
        pts.append(f'<c:pt idx="{i}"><c:v>{v_str}</c:v></c:pt>')
    fmt = f"<c:formatCode>{format_code}</c:formatCode>" if format_code else ""
    return f'<{tag}>{fmt}<c:ptCount val="{len(values)}"/>{"".join(pts)}</{tag}>'


def process_ref_block(block_xml, wb, grow=True, report=None, no_grow_sheets=frozenset()):
    """
    Recibe el contenido de un <c:cat> o <c:val> (incluye <c:numRef>/<c:strRef>),
    localiza su <c:f>, relee los datos del Excel (extendiendo el rango si
    corresponde) y devuelve el bloque reescrito.

    `no_grow_sheets`: nombres de hoja donde NUNCA se crece el rango más allá
    de lo ya declarado (sí se sigue recortando si el final tiene celdas
    vacías). Para hojas de cálculo/resumen de uso mixto, donde "la celda
    de al lado" puede pertenecer a otro bloque de datos sin relación con el
    gráfico, en vez de ser el siguiente dato real de la serie.
    """
    f_match = re.search(r"<c:f>([^<]*)</c:f>", block_xml)
    if not f_match:
        return block_xml  # literal (c:numLit/c:strLit) sin vínculo a celdas: no se toca
    ref = f_match.group(1)
    try:
        sheet, c1, r1, c2, r2 = parse_ref(ref)
    except RefParseError as e:
        if report is not None:
            report.append(f"  ! {e}")
        return block_xml

    if sheet not in wb.sheetnames:
        if report is not None:
            report.append(f"  ! hoja '{sheet}' no existe en el Excel, se deja igual")
        return block_xml
    ws = wb[sheet]

    is_num = "<c:numRef>" in block_xml or "<c:numCache>" in block_xml
    cache_tag = "c:numCache" if is_num else "c:strCache"
    fmt_match = re.search(r"<c:formatCode>([^<]*)</c:formatCode>", block_xml)
    format_code = fmt_match.group(1) if fmt_match else None
    is_date = bool(format_code and ("m" in format_code.lower() and "y" in format_code.lower()))

    new_c1, new_r1, new_c2, new_r2 = (c1, r1, c2, r2)
    if grow:
        max_scan = 0 if sheet in no_grow_sheets else MAX_SCAN
        new_c1, new_r1, new_c2, new_r2 = scan_extent(ws, sheet, c1, r1, c2, r2, max_scan=max_scan)

    values = read_values(ws, new_c1, new_r1, new_c2, new_r2)
    new_ref = format_ref(sheet, new_c1, new_r1, new_c2, new_r2)
    new_cache = build_cache_xml(cache_tag, format_code, values, is_date=is_date)

    if report is not None and (new_ref != ref or True):
        n = sum(1 for v in values if v is not None)
        growth = f" (antes {ref})" if new_ref != ref else ""
        report.append(f"  {new_ref}{growth}: {n} puntos")

    new_block = re.sub(r"<c:f>[^<]*</c:f>", f"<c:f>{new_ref}</c:f>", block_xml, count=1)
    cache_re = re.compile(r"<c:numCache>.*?</c:numCache>|<c:strCache>.*?</c:strCache>", re.S)
    if cache_re.search(new_block):
        new_block = cache_re.sub(lambda m: new_cache, new_block, count=1)
    else:
        # no había caché (poco común): insertarla justo después de </c:f>
        new_block = new_block.replace(f"<c:f>{new_ref}</c:f>", f"<c:f>{new_ref}</c:f>{new_cache}", 1)
    return new_block


def process_series_block(ser_xml, wb, report=None, no_grow_sheets=frozenset()):
    name_match = re.search(r"<c:tx>.*?<c:v>([^<]*)</c:v>", ser_xml, re.S)
    sname = name_match.group(1) if name_match else "(sin nombre)"
    if report is not None:
        report.append(f"- serie: {sname}")

    def repl_cat(m):
        return "<c:cat>" + process_ref_block(m.group(1), wb, grow=True, report=report, no_grow_sheets=no_grow_sheets) + "</c:cat>"

    def repl_val(m):
        return "<c:val>" + process_ref_block(m.group(1), wb, grow=True, report=report, no_grow_sheets=no_grow_sheets) + "</c:val>"

    ser_xml = re.sub(r"<c:cat>(.*?)</c:cat>", repl_cat, ser_xml, count=1, flags=re.S)
    ser_xml = re.sub(r"<c:val>(.*?)</c:val>", repl_val, ser_xml, count=1, flags=re.S)
    return ser_xml


def sync_chart_xml(xml_text, wb, report=None, no_grow_sheets=frozenset()):
    ser_re = re.compile(r"<c:ser>.*?</c:ser>", re.S)
    spans = [m.span() for m in ser_re.finditer(xml_text)]
    if not spans:
        return xml_text
    pieces = [xml_text[: spans[0][0]]]
    for i, (start, end) in enumerate(spans):
        new_block = process_series_block(xml_text[start:end], wb, report=report, no_grow_sheets=no_grow_sheets)
        pieces.append(new_block)
        next_start = spans[i + 1][0] if i + 1 < len(spans) else len(xml_text)
        pieces.append(xml_text[end:next_start])
    return "".join(pieces)


def sync_charts_in_zip(src_path, out_path, wb, chart_prefix, chart_filter=None, verbose=True, no_grow_sheets=frozenset()):
    """
    Motor genérico: reescribe los chartN.xml de CUALQUIER paquete OOXML
    (un .pptx o un .xlsx; ambos son zips con el mismo esquema de gráficos)
    usando los valores actuales de `wb` (un Workbook de openpyxl abierto con
    data_only=True). `chart_prefix` es "ppt/charts/" o "xl/charts/".
    """
    src_path, out_path = str(src_path), str(out_path)
    same_file = Path(src_path).resolve() == Path(out_path).resolve()
    tmp_out = out_path + ".tmp"
    with zipfile.ZipFile(src_path, "r") as zin:
        chart_names = sorted(
            n for n in zin.namelist()
            if re.match(re.escape(chart_prefix) + r"chart\d+\.xml$", n)
        )
        if chart_filter:
            chart_names = [n for n in chart_names if any(f in n for f in chart_filter)]

        updates = {}
        for name in chart_names:
            xml_text = zin.read(name).decode("utf-8")
            report = [] if verbose else None
            if verbose:
                print(f"\n== {name} ==")
            new_xml = sync_chart_xml(xml_text, wb, report=report, no_grow_sheets=no_grow_sheets)
            if verbose:
                for line in report:
                    print(line)
            updates[name] = new_xml

        # reescribe el zip completo, sustituyendo sólo los chartN.xml tocados
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in updates:
                    data = updates[item.filename].encode("utf-8")
                zout.writestr(item, data)
    shutil.move(tmp_out, out_path)
    return list(updates.keys())


def sync_pptx(pptx_path, xlsx_path, out_path, chart_filter=None, verbose=True, no_grow_sheets=frozenset()):
    """Actualiza los gráficos de un .pptx a partir de un Excel EXTERNO."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    return sync_charts_in_zip(pptx_path, out_path, wb, "ppt/charts/", chart_filter=chart_filter, verbose=verbose, no_grow_sheets=no_grow_sheets)


def sync_xlsx_own_charts(xlsx_path, out_path, chart_filter=None, verbose=True, no_grow_sheets=frozenset()):
    """
    Actualiza los gráficos NATIVOS de un .xlsx (los que se ven al abrirlo en
    Excel) a partir de sus propias hojas de datos. Es el mismo mecanismo que
    sync_pptx: cada gráfico ya trae una fórmula de rango (p.ej.
    "Gasto!$B$2:$B$9"); esto relee esas celdas, extiende el rango si hay
    filas/columnas nuevas y reescribe la caché del gráfico. Como la caché y
    el dato en sí viven en el mismo archivo, se lee con el propio
    openpyxl(data_only=True) sobre xlsx_path antes de tocar el zip.
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    return sync_charts_in_zip(xlsx_path, out_path, wb, "xl/charts/", chart_filter=chart_filter, verbose=verbose, no_grow_sheets=no_grow_sheets)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", required=True, help="Excel fuente (el que alimenta los gráficos)")
    ap.add_argument("--pptx", required=True, help="Presentación .pptx a actualizar")
    ap.add_argument("--out", help="Ruta de salida del pptx (por defecto: <pptx>.synced.pptx)")
    ap.add_argument("--chart", action="append", help="Limitar a chartN.xml específicos, p.ej. --chart chart1 --chart chart7")
    ap.add_argument(
        "--skip-excel-chart",
        action="store_true",
        help="No tocar los gráficos nativos del propio Excel (por defecto SÍ se actualizan, en el mismo archivo --xlsx).",
    )
    ap.add_argument(
        "--no-grow-sheet",
        action="append",
        default=[],
        help="Nombre de hoja donde nunca crecer un rango más allá de lo ya declarado "
             "(sí se sigue recortando si el final tiene celdas vacías). Repetible. "
             "Útil para hojas de uso mixto donde la celda de al lado puede pertenecer "
             "a otro bloque de datos sin relación con el gráfico.",
    )
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    pptx_path = Path(args.pptx)
    xlsx_path = Path(args.xlsx)
    out_path = Path(args.out) if args.out else pptx_path.with_suffix(".synced.pptx")
    verbose = not args.quiet
    no_grow_sheets = set(args.no_grow_sheet)

    touched = sync_pptx(pptx_path, xlsx_path, out_path, chart_filter=args.chart, verbose=verbose, no_grow_sheets=no_grow_sheets)
    print(f"\nListo. {len(touched)} gráfico(s) del pptx actualizados -> {out_path}")

    if not args.skip_excel_chart:
        touched_xlsx = sync_xlsx_own_charts(xlsx_path, xlsx_path, chart_filter=args.chart, verbose=verbose, no_grow_sheets=no_grow_sheets)
        print(f"\nListo. {len(touched_xlsx)} gráfico(s) del propio Excel actualizados -> {xlsx_path}")


if __name__ == "__main__":
    main()
