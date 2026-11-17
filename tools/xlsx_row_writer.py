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

from openpyxl.utils import column_index_from_string

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


def replace_cell(xlsx_path, sheet_name, cell_ref, kind, value, style=None):
    """
    Reemplaza UNA celda ya existente (por ejemplo, pasar de la fórmula
    supuesta "+E163*(1+3%)" al dato real publicado por el Banco Central),
    editando sólo esa celda en el XML de la hoja. Si no se pasa `style`, se
    conserva el estilo que ya tenía la celda.
    """
    xlsx_path = str(xlsx_path)
    tmp_path = xlsx_path + ".celltmp"
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        sheet_file = _find_sheet_file(zin, sheet_name)
        sheet_xml = zin.read(sheet_file).decode("utf-8")

        m = re.search(r'<c r="' + re.escape(cell_ref) + r'"(?:\s+s="(\d+)")?[^>]*(?:/>|>.*?</c>)', sheet_xml)
        if not m:
            raise ValueError(f"No encontré la celda {cell_ref} en la hoja {sheet_name!r}")
        if style is None:
            style = m.group(1)

        new_cell_xml = _cell_xml(cell_ref, kind, value, style)
        new_sheet_xml = sheet_xml[: m.start()] + new_cell_xml + sheet_xml[m.end():]

        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == sheet_file:
                    data = new_sheet_xml.encode("utf-8")
                zout.writestr(item, data)
    shutil.move(tmp_path, xlsx_path)


def update_formula_and_value(xlsx_path, sheet_name, cell_ref, new_formula_text, new_value):
    """
    Cambia TANTO el texto de la fórmula como su valor cacheado de una celda
    que ya existe, preservando cualquier atributo del <f> (p. ej.
    t="shared" ref="..." si="N", para no romper el grupo de fórmula
    compartida del que es "maestra"). Úsala cuando la fórmula en sí debe
    cambiar (p. ej. pasar de proyectar con el forward a leer el dato real
    recién publicado), a diferencia de update_cached_value (que sólo
    refresca el resultado de una fórmula que no cambia).
    """
    xlsx_path = str(xlsx_path)
    tmp_path = xlsx_path + ".fvtmp"
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        sheet_file = _find_sheet_file(zin, sheet_name)
        sheet_xml = zin.read(sheet_file).decode("utf-8")

        m = re.search(r'<c r="' + re.escape(cell_ref) + r'"[^>]*>.*?</c>', sheet_xml)
        if not m:
            raise ValueError(f"No encontré la celda {cell_ref} en la hoja {sheet_name!r}")
        cell_xml = m.group(0)
        if "<f" not in cell_xml:
            raise ValueError(f"{cell_ref} no tiene fórmula; usa replace_cell para ponerle una.")
        new_cell_xml = re.sub(r"(<f[^>]*>)[^<]*(</f>)", lambda mm: mm.group(1) + _xml_escape(new_formula_text) + mm.group(2), cell_xml, count=1)
        if "<v>" in new_cell_xml:
            new_cell_xml = re.sub(r"<v>[^<]*</v>", f"<v>{new_value!r}</v>", new_cell_xml, count=1)
        else:
            new_cell_xml = new_cell_xml.replace("</f>", f"</f><v>{new_value!r}</v>", 1)

        new_sheet_xml = sheet_xml[: m.start()] + new_cell_xml + sheet_xml[m.end():]

        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == sheet_file:
                    data = new_sheet_xml.encode("utf-8")
                zout.writestr(item, data)
    shutil.move(tmp_path, xlsx_path)


def insert_cells_into_row(xlsx_path, sheet_name, row_num, cells):
    """
    Agrega celdas a una fila que YA EXISTE pero no las tiene todas (p. ej.
    una fila "plantilla" que ya trae la fecha pero no los 15 valores, a la
    espera del próximo dato). A diferencia de append_row (que crea la fila
    completa), esto respeta lo que la fila ya tenía y sólo agrega lo que
    falta, en el orden de columna correcto (obligatorio en OOXML).

    cells: misma forma que en append_row — lista de
    (col_letter, kind, value, style_or_None). Si una columna ya tiene celda
    en esa fila, se dejará como estaba (no se pisa).
    """
    xlsx_path = str(xlsx_path)
    tmp_path = xlsx_path + ".instmp"
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        sheet_file = _find_sheet_file(zin, sheet_name)
        sheet_xml = zin.read(sheet_file).decode("utf-8")

        row_m = re.search(r'<row r="' + str(row_num) + r'"[^>]*>(.*?)</row>', sheet_xml, re.S)
        if not row_m:
            raise ValueError(f"La fila {row_num} de {sheet_name!r} no existe; usa append_row.")
        row_inner = row_m.group(1)

        existing = {}
        for cm in re.finditer(r'<c r="([A-Z]+)' + str(row_num) + r'"[^>]*(?:/>|>.*?</c>)', row_inner):
            existing[cm.group(1)] = cm.span()

        new_cells = {}
        for col, kind, value, style in cells:
            if col in existing:
                continue
            if style is None:
                style = _style_above(sheet_xml, col, row_num - 1)
            new_cells[col] = _cell_xml(f"{col}{row_num}", kind, value, style)

        if not new_cells:
            raise ValueError(f"Todas las columnas pedidas ya existían en la fila {row_num} de {sheet_name!r}.")

        # arma la fila completa en orden de columna
        all_cols = sorted(set(existing) | set(new_cells), key=column_index_from_string)
        pieces = []
        for col in all_cols:
            if col in new_cells:
                pieces.append(new_cells[col])
            else:
                start, end = existing[col]
                pieces.append(row_inner[start:end])
        new_row_inner = "".join(pieces)
        new_sheet_xml = sheet_xml[: row_m.start(1)] + new_row_inner + sheet_xml[row_m.end(1):]

        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == sheet_file:
                    data = new_sheet_xml.encode("utf-8")
                zout.writestr(item, data)
    shutil.move(tmp_path, xlsx_path)


def update_cached_value(xlsx_path, sheet_name, cell_ref, new_value):
    """
    Actualiza sólo el valor cacheado (<v>) de una celda que ya tiene una
    fórmula, dejando la fórmula (<f>...) intacta tal como está. Se usa
    cuando cambia un valor del que depende la fórmula (p. ej. el PIB
    nominal real reemplaza al supuesto) y hay que refrescar el resultado
    sin tocar cómo se calcula.
    """
    xlsx_path = str(xlsx_path)
    tmp_path = xlsx_path + ".cvtmp"
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        sheet_file = _find_sheet_file(zin, sheet_name)
        sheet_xml = zin.read(sheet_file).decode("utf-8")

        m = re.search(r'<c r="' + re.escape(cell_ref) + r'"[^>]*>.*?</c>', sheet_xml)
        if not m:
            raise ValueError(f"No encontré la celda {cell_ref} en la hoja {sheet_name!r}")
        cell_xml = m.group(0)
        if "<v>" in cell_xml:
            new_cell_xml = re.sub(r"<v>[^<]*</v>", f"<v>{new_value!r}</v>", cell_xml, count=1)
        else:
            new_cell_xml = cell_xml.replace("</f>", f"</f><v>{new_value!r}</v>", 1)

        new_sheet_xml = sheet_xml[: m.start()] + new_cell_xml + sheet_xml[m.end():]

        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == sheet_file:
                    data = new_sheet_xml.encode("utf-8")
                zout.writestr(item, data)
    shutil.move(tmp_path, xlsx_path)


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
