import re

_CELL_RE = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d{1,7})")


def shift_formula_rows(formula, delta):
    """
    Dado el texto de una fórmula (tal como la escribirías en Excel, con o
    sin el '=' inicial), devuelve la misma fórmula con cada referencia de
    fila RELATIVA desplazada en `delta`. Las referencias de fila absolutas
    (con '$' antes del número de fila, ej. $B$7) no se tocan — así es como
    Excel copia una fórmula una fila hacia abajo.
    """
    def repl(m):
        dollar_col, col, dollar_row, row = m.groups()
        if dollar_row == "$":
            return m.group(0)
        return f"{dollar_col}{col}{dollar_row}{int(row) + delta}"
    return _CELL_RE.sub(repl, formula)


def copy_formula_down(ws, from_row, to_row, col_letter):
    """
    Copia la fórmula de ws[col_letter][from_row] a ws[col_letter][to_row],
    desplazando las referencias relativas de fila. Si la celda de origen no
    es una fórmula, no hace nada (deja to_row como está).
    """
    src = ws[f"{col_letter}{from_row}"]
    if not (isinstance(src.value, str) and src.value.startswith("=")):
        return False
    delta = to_row - from_row
    ws[f"{col_letter}{to_row}"] = shift_formula_rows(src.value, delta)
    return True
