"""
Utilidades auxiliares para trabajar con el Excel fuente de los gráficos.

Aviso importante: varias columnas de fecha en "Graficos 11112026.xlsx" usan
fórmulas arrastradas (p.ej. "=A3+7" copiada hacia abajo como fórmula
compartida). openpyxl puede LEER el valor cacheado de esas fórmulas
(load_workbook(..., data_only=True)), pero si luego ABRES y GUARDAS ese mismo
archivo con openpyxl (no con Excel) sin tocar esas celdas, openpyxl no
reevalúa las fórmulas y el valor cacheado se pierde (la celda queda en
blanco). pptx_chart_sync.py sólo LEE el Excel, nunca lo guarda, así que no
tiene este problema. Pero si en el futuro automatizas también la escritura de
filas nuevas (por ejemplo, para insertar datos del Banco Central), usa
`flatten_formulas_to_values` primero para "aplanar" el archivo a valores
literales antes de guardarlo con openpyxl.
"""
import openpyxl


def flatten_formulas_to_values(src_path, dst_path):
    """
    Crea una copia de src_path en dst_path donde cada celda queda con su
    último valor calculado (literal), sin fórmulas. Útil antes de que un
    script (no Excel) vuelva a guardar el archivo, para no perder los
    valores cacheados de columnas con fórmulas arrastradas.
    """
    wb_vals = openpyxl.load_workbook(src_path, data_only=True)
    wb_out = openpyxl.load_workbook(src_path, data_only=False)
    for sheet_name in wb_out.sheetnames:
        ws_vals = wb_vals[sheet_name]
        ws_out = wb_out[sheet_name]
        for row in ws_out.iter_rows():
            for cell in row:
                v = ws_vals[cell.coordinate].value
                if cell.value is not None or v is not None:
                    cell.value = v
    wb_out.save(dst_path)
    return dst_path
