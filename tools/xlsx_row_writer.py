"""
Agrega una fila nueva a una hoja de Graficos 11112026.xlsx editando
directamente el XML interno (como pptx_chart_sync.py hace con los
gráficos), en vez de abrir y volver a guardar el libro con openpyxl.

Por qué: las filas de este Excel tienen columnas de apoyo con fórmulas
compartidas ("shared formulas", <f t="shared" si="N"/>) fuera de las que
alimentan los gráficos. openpyxl puede LEER esas fórmulas sin problema,
pero si abres el libro completo y lo vuelves a guardar con openpyxl —aunque
sólo hayas tocado 2 celdas— no reevalúa esas fórmulas compartidas y el
valor cacheado de TODAS ellas, en CUALQUIER hoja, se pierde (queda en
blanco). Esto se confirmó en pruebas: un guardado así dejó sin datos la
columna de fechas de la hoja "Combustibles", que nadie tocó.

Esta función evita el problema por completo: nunca abre el libro con
openpyxl en modo escritura. Sólo inserta un <row> nuevo con las celdas que
le pasas (las que alimentan un gráfico); todo lo demás del archivo —
incluidas las columnas de apoyo que no tocamos— queda exactamente igual.
"""
import re
import shutil
import zipfile
from datetime import date, datetime
from pathlib import Path

EXCEL_EPOCH = date(1899, 12, 30)


def to_serial(d):
    if isinstance(d, datetime):
        d = d.date()
    return (d - EXCEL_EPOCH).days


def _xml_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _find_sheet_file(zf, sheet_name):
    wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    m = re.search(r'<sheet name="' + re.escape(sheet_name) + r'"[^>]*r:id="(rId\d+)"', wb_xml)
    if not m:
        raise ValueError(f"No encontré la hoja {sheet_name!r} en xl/workbook.xml")
    rid = m.group(1)
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m2 = re.search(r'<Relationship Id="' + rid + r'"[^>]*Target="([^"]+)"', rels)
    return "xl/" + m2.group(1)


def _style_above(sheet_xml, col, row_above):
    m = re.search(r'<c r="' + col + str(row_above) + r'"(?:\s+s="(\d+)")?', sheet_xml)
    return m.group(1) if (m and m.group(1)) else None


def _cell_xml(ref, kind, value, style):
    s_attr = f' s="{style}"' if style else ""
    if kind == "num":
        return f'<c r="{ref}"{s_attr}><v>{value!r}</v></c>'
    if kind == "date":
        return f'<c r="{ref}"{s_attr}><v>{to_serial(value)}</v></c>'
    if kind == "str":
        return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t>{_xml_escape(str(value))}</t></is></c>'
    if kind == "formula":
        formula_text, cached = value
        cached_xml = "" if cached is None else f"<v>{cached!r}</v>"
        return f'<c r="{ref}"{s_attr}><f>{_xml_escape(formula_text)}</f>{cached_xml}</c>'
    raise ValueError(f"tipo de celda desconocido: {kind}")


def append_row(xlsx_path, sheet_name, row_num, cells):
    """
    cells: lista de (col_letter, kind, value, style_or_None)
      kind: "num" (value=float), "date" (value=date/datetime),
            "str" (value=texto), "formula" (value=(formula_sin_igual, valor_cacheado))
      style: índice de estilo a usar; si es None, se copia el de la misma
             columna en la fila anterior (row_num - 1).
    Escribe in-place sobre xlsx_path.
    """
    xlsx_path = str(xlsx_path)
    tmp_path = xlsx_path + ".rowtmp"
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        sheet_file = _find_sheet_file(zin, sheet_name)
        sheet_xml = zin.read(sheet_file).decode("utf-8")

        if re.search(r'<row r="' + str(row_num) + r'"[ >]', sheet_xml):
            raise ValueError(f"La fila {row_num} de {sheet_name!r} ya existe; no se puede agregar de nuevo.")

        parts = []
        for col, kind, value, style in cells:
            if style is None:
                style = _style_above(sheet_xml, col, row_num - 1)
            parts.append(_cell_xml(f"{col}{row_num}", kind, value, style))

        spans_m = re.search(r'<row r="' + str(row_num - 1) + r'"[^>]*\bspans="([^"]+)"', sheet_xml)
        spans_attr = f' spans="{spans_m.group(1)}"' if spans_m else ""
        new_row_xml = f'<row r="{row_num}"{spans_attr}>' + "".join(parts) + "</row>"

        new_sheet_xml = sheet_xml.replace("</sheetData>", new_row_xml + "</sheetData>", 1)

        def bump_dim(m):
            start, end = m.group(1).split(":")
            end_col = re.match(r"[A-Z]+", end).group(0)
            end_row = int(re.search(r"\d+", end).group(0))
            if row_num > end_row:
                end = f"{end_col}{row_num}"
            return f'<dimension ref="{start}:{end}"/>'

        new_sheet_xml = re.sub(r'<dimension ref="([^"]+)"/>', bump_dim, new_sheet_xml, count=1)

        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == sheet_file:
                    data = new_sheet_xml.encode("utf-8")
                zout.writestr(item, data)
    shutil.move(tmp_path, xlsx_path)
