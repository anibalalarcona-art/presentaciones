#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Trae el Tipo de Cambio Real (REER) del BIS y, si hay un mes nuevo, actualiza
en cascada Gráficos_TC.xlsx (hojas "TCR" y "TCR (ajustado)", conservando sus
fórmulas) y después sincroniza los gráficos de Láminas_TC.pptx y del propio
Excel con pptx_chart_sync.py.

Qué toca, en orden, cuando aparece un mes nuevo (p. ej. octubre):
  1. Hoja "TCR": llena las 15 columnas (B:P) de la fila ya pre-armada con la
     fecha (venían con la fecha puesta y el valor en blanco, a la espera).
     Extiende las fórmulas "Prom. 5a"/"Prom. 10a" (AVERAGE) para que incluyan
     la fila nueva.
  2. Hoja "TCR (ajustado)": esa misma fila tenía una fórmula que PROYECTABA
     el valor con el forward de tipo de cambio (p. ej. "+B145*TCN!B208/
     TCN!B207") mientras no existía el dato real. La reemplaza por la misma
     fórmula que ya usan los meses reales (p. ej. "(100/TCR!B146)*100"),
     tomada de la fila anterior y desplazada una fila — no se inventa el
     patrón, se copia el que el propio archivo ya usa.
  3. Recalcula en cascada las filas de proyección siguientes (que ahora
     parten de un valor real en vez de uno proyectado), los promedios 5a/10a
     de esa hoja, y "TCR estimado <mes>" (que pasa a apuntar al siguiente
     mes, el que todavía no tiene dato real).
  4. Reordena y reescribe la tabla de ranking (filas 156-170: los 15 países
     ordenados de mayor a menor TCR, la que alimenta los gráficos de las
     láminas 8 y 9).
  5. Llama a pptx_chart_sync.py (con --no-grow-sheet Resumen, para no volver
     a arrastrar la columna "2027" que ya se sacó a mano).

Uso:
    python3 tools/update_tc_from_apis.py --xlsx "Gráficos_TC.xlsx" \
        --pptx "Láminas_TC.pptx" --apply

    Sin --apply, sólo muestra qué haría (dry run).
"""
import argparse
import datetime
import re
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
import bis_client
import xlsx_row_writer as rw
import pptx_chart_sync
from formula_utils import shift_formula_rows

MESES_ABR = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
COLS = [chr(ord("B") + i) for i in range(15)]  # B..P


def month_label(d):
    return f"{MESES_ABR[d.month - 1]}-{str(d.year)[2:]}"


def last_real_row(ws_values, start_row=5):
    """Última fila con dato en columna B (no sólo fecha en A)."""
    r = start_row
    last = None
    while ws_values[f"A{r}"].value is not None:
        if ws_values[f"B{r}"].value is not None:
            last = r
        r += 1
    return last


def find_target_row(ws_values, after_row, target_date):
    """Busca, a partir de after_row+1, la fila cuya fecha (col A) sea target_date."""
    r = after_row + 1
    while ws_values[f"A{r}"].value is not None:
        d = ws_values[f"A{r}"].value
        if hasattr(d, "date"):
            d = d.date()
        if d == target_date:
            return r
        r += 1
    return None


def bump_average_end(formula_text, new_end_row):
    """'+AVERAGE(B86:B145)' con new_end_row=146 -> '+AVERAGE(B86:B146)' (no baja el final si ya era mayor)."""
    m = re.search(r"AVERAGE\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)", formula_text)
    if not m:
        return formula_text, None
    c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
    new_r2 = max(r2, new_end_row)
    new_formula = formula_text[: m.start()] + f"AVERAGE({c1}{r1}:{c2}{new_r2})" + formula_text[m.end():]
    return new_formula, (c1, r1, c2, new_r2)


def avg_range_value(ws_values, col, r1, r2):
    vals = [ws_values[f"{col}{r}"].value for r in range(r1, r2 + 1)]
    vals = [v for v in vals if v is not None]
    return sum(vals) / len(vals) if vals else None


def update_tcr_sheet(xlsx_path, target_row, new_values_by_col, log):
    """Paso 1: llena TCR!<target_row> y extiende Prom. 5a / Prom. 10a."""
    cells = [(col, "num", new_values_by_col[col], None) for col in COLS]
    rw.insert_cells_into_row(xlsx_path, "TCR", target_row, cells)
    log.append(f"[TCR] fila {target_row} llenada con los 15 países.")

    wbf = openpyxl.load_workbook(xlsx_path, data_only=False)
    wbv = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws_f, ws_v = wbf["TCR"], wbv["TCR"]

    for avg_row in (150, 151):  # Prom. 5a, Prom. 10a
        for col in COLS:
            formula = ws_f[f"{col}{avg_row}"].value
            if not (isinstance(formula, str) and formula.startswith("=")):
                continue
            new_formula, parsed = bump_average_end(formula[1:], target_row)
            if parsed is None:
                continue
            c1, r1, c2, r2 = parsed
            new_value = avg_range_value(ws_v, col, r1, r2)
            if new_value is None:
                continue
            rw.update_formula_and_value(xlsx_path, "TCR", f"{col}{avg_row}", new_formula, new_value)
    log.append(f"[TCR] Prom. 5a / Prom. 10a extendidos hasta la fila {target_row}.")


_INVERTED_RE = re.compile(r"^\(100/TCR!(\w+\d+)\)\*100$")
_DIRECT_RE = re.compile(r"^\+TCR!(\w+\d+)$")
_PROJECTION_RE = re.compile(r"^\+(\w+)(\d+)\*TCN!(\w+)(\d+)/TCN!(\w+)(\d+)$")


def eval_real_transform(formula_text, tcr_value):
    m = _INVERTED_RE.match(formula_text)
    if m:
        return (100.0 / tcr_value) * 100.0
    m = _DIRECT_RE.match(formula_text)
    if m:
        return tcr_value
    raise ValueError(f"patrón de fórmula no reconocido en 'ajustado': {formula_text!r}")


def update_ajustado_sheet(xlsx_path, target_row, log):
    """Pasos 2-3: flip la fila del mes nuevo y recalcula la cascada de proyección."""
    wbf = openpyxl.load_workbook(xlsx_path, data_only=False)
    wbv = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws_f, ws_v = wbf["TCR (ajustado)"], wbv["TCR (ajustado)"]
    ws_tcr_v = wbv["TCR"]

    chain_values = {}  # row -> {col: value}, valores "ajustado" ya recalculados en esta corrida
    for col in COLS:
        prev_formula = ws_f[f"{col}{target_row - 1}"].value
        if not (isinstance(prev_formula, str) and prev_formula.startswith("=")):
            raise ValueError(f"{col}{target_row - 1} de 'TCR (ajustado)' no es una fórmula; no puedo copiarla.")
        new_formula = shift_formula_rows(prev_formula[1:], 1)
        tcr_value = ws_tcr_v[f"{col}{target_row}"].value
        new_value = eval_real_transform(new_formula, tcr_value)
        rw.update_formula_and_value(xlsx_path, "TCR (ajustado)", f"{col}{target_row}", new_formula, new_value)
        chain_values.setdefault(target_row, {})[col] = new_value
    log.append(f"[TCR ajustado] fila {target_row} pasó de proyección a dato real.")

    # recalcula las filas de proyección siguientes (si existen), encadenadas
    r = target_row + 1
    while True:
        existing_formula = ws_f[f"B{r}"].value
        if not (isinstance(existing_formula, str) and existing_formula.startswith("=")):
            break
        any_col = False
        for col in COLS:
            formula = ws_f[f"{col}{r}"].value
            m = _PROJECTION_RE.match(formula[1:]) if isinstance(formula, str) else None
            if not m:
                continue
            prev_col, prev_row, tcn_col1, tcn_row1, tcn_col2, tcn_row2 = m.groups()
            prev_row = int(prev_row)
            prev_val = chain_values.get(prev_row, {}).get(prev_col)
            if prev_val is None:
                prev_val = ws_v[f"{prev_col}{prev_row}"].value
            tcn1 = ws_v[f"{tcn_col1}{tcn_row1}"].value
            tcn2 = ws_v[f"{tcn_col2}{tcn_row2}"].value
            if prev_val is None or tcn1 is None or not tcn2:
                continue  # deja el error (#DIV/0! etc) tal como estaba
            new_value = prev_val * tcn1 / tcn2
            rw.update_cached_value(xlsx_path, "TCR (ajustado)", f"{col}{r}", new_value)
            chain_values.setdefault(r, {})[col] = new_value
            any_col = True
        if any_col:
            log.append(f"[TCR ajustado] fila {r} (proyección) recalculada con el nuevo dato base.")
        r += 1

    return chain_values


def update_ajustado_summary(xlsx_path, target_row, chain_values, log):
    """Paso 3 (cont.): Prom. 5a / Prom. 10a / 'TCR estimado <mes>' de 'TCR (ajustado)'."""
    wbf = openpyxl.load_workbook(xlsx_path, data_only=False)
    wbv = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws_f, ws_v = wbf["TCR (ajustado)"], wbv["TCR (ajustado)"]

    summary_row = None
    for r in range(target_row + 1, target_row + 10):
        if isinstance(ws_f[f"A{r}"].value, str) and ws_f[f"A{r}"].value.startswith("Prom"):
            summary_row = r
            break
    if summary_row is None:
        log.append("[TCR ajustado] no encontré las filas de resumen (Prom. 5a/10a); no se tocaron.")
        return {}

    row_5a, row_10a, row_estimado = summary_row, summary_row + 1, summary_row + 2

    # "TCR estimado <mes>" (y el rango de los promedios, más abajo) siempre
    # incluye la fila SIGUIENTE a la recién confirmada: la próxima que aún
    # no tiene dato real (proyección) — o la propia target_row si no hay
    # ninguna fila de proyección después.
    next_row = target_row + 1
    has_next = isinstance(ws_f[f"B{next_row}"].value, str)
    estimado_row_ref = next_row if has_next else target_row
    next_date = ws_v[f"A{estimado_row_ref}"].value
    next_label = f"TCR estimado {month_label(next_date.date() if hasattr(next_date,'date') else next_date)}"

    avgs = {}
    for avg_row in (row_5a, row_10a):
        for col in COLS:
            formula = ws_f[f"{col}{avg_row}"].value
            if not (isinstance(formula, str) and formula.startswith("=")):
                continue
            new_formula, parsed = bump_average_end(formula[1:], estimado_row_ref)
            if parsed is None:
                continue
            c1, r1, c2, r2 = parsed
            vals = []
            for rr in range(r1, r2 + 1):
                v = chain_values.get(rr, {}).get(col)
                if v is None:
                    v = ws_v[f"{col}{rr}"].value
                if v is not None:
                    vals.append(v)
            if not vals:
                continue
            new_value = sum(vals) / len(vals)
            rw.update_formula_and_value(xlsx_path, "TCR (ajustado)", f"{col}{avg_row}", new_formula, new_value)
            avgs.setdefault(avg_row, {})[col] = new_value

    rw.replace_cell(xlsx_path, "TCR (ajustado)", f"A{row_estimado}", "str", next_label)
    estimado_vals = {}
    for col in COLS:
        v = chain_values.get(estimado_row_ref, {}).get(col)
        if v is None:
            v = ws_v[f"{col}{estimado_row_ref}"].value
        new_formula = f"+{col}{estimado_row_ref}"
        rw.update_formula_and_value(xlsx_path, "TCR (ajustado)", f"{col}{row_estimado}", new_formula, v)
        estimado_vals[col] = v
    log.append(f"[TCR ajustado] resumen actualizado: '{next_label}' (fila {estimado_row_ref}).")

    # la tabla de ranking tiene su propio encabezado más abajo (columnas D/H,
    # no A): "Prom. 5a" / "TCR estimado <mes>" otra vez, en formato angosto.
    header_row = None
    for r in range(row_estimado + 1, row_estimado + 10):
        d_val = ws_f[f"D{r}"].value
        if isinstance(d_val, str) and d_val.startswith("Prom"):
            header_row = r
            break
    if header_row is None:
        log.append("[TCR ajustado] no encontré el encabezado de la tabla de ranking; no se reordenó.")
        header_row = None

    return {
        "avg5": avgs.get(row_5a, {}),
        "avg10": avgs.get(row_10a, {}),
        "estimado": estimado_vals,
        "label": next_label,
        "header_row": header_row,
    }


def update_ranking_table(xlsx_path, summary, log):
    """Paso 4: reordena la tabla de 15 países por TCR, dos bloques (5a y 10a)."""
    names = {code: name for name, code in bis_client.COUNTRIES}
    # columnas B:P de 'ajustado' están en el mismo orden que bis_client.COUNTRIES
    col_to_name = dict(zip(COLS, [name for name, _ in bis_client.COUNTRIES]))

    avg5, avg10, estimado = summary["avg5"], summary["avg10"], summary["estimado"]
    if not avg5 or not avg10 or summary.get("header_row") is None:
        log.append("[Ranking] faltaron promedios o encabezado; no se reordenó la tabla.")
        return

    header_row = summary["header_row"]
    ranking_start = header_row + 1

    rw.replace_cell(xlsx_path, "TCR (ajustado)", f"E{header_row}", "str", summary["label"])
    rw.replace_cell(xlsx_path, "TCR (ajustado)", f"I{header_row}", "str", summary["label"])

    rows5 = sorted(col_to_name.items(), key=lambda kv: avg5.get(kv[0], float("-inf")), reverse=True)
    rows10 = sorted(col_to_name.items(), key=lambda kv: avg10.get(kv[0], float("-inf")), reverse=True)

    for i, (col, name) in enumerate(rows5):
        r = ranking_start + i
        rw.replace_cell(xlsx_path, "TCR (ajustado)", f"C{r}", "str", name)
        rw.replace_cell(xlsx_path, "TCR (ajustado)", f"D{r}", "num", avg5[col])
        rw.replace_cell(xlsx_path, "TCR (ajustado)", f"E{r}", "num", estimado.get(col, avg5[col]))

    for i, (col, name) in enumerate(rows10):
        r = ranking_start + i
        rw.replace_cell(xlsx_path, "TCR (ajustado)", f"G{r}", "str", name)
        rw.replace_cell(xlsx_path, "TCR (ajustado)", f"H{r}", "num", avg10[col])
        rw.replace_cell(xlsx_path, "TCR (ajustado)", f"I{r}", "num", estimado.get(col, avg10[col]))

    log.append(f"[Ranking] tabla reordenada ({len(rows5)} países) en filas {ranking_start}-{ranking_start+14}.")


def run(xlsx_path, pptx_path, log, apply):
    wbv = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws_v = wbv["TCR"]
    last_row = last_real_row(ws_v)
    last_date = ws_v[f"A{last_row}"].value
    last_date = last_date.date() if hasattr(last_date, "date") else last_date
    log.append(f"[TCR] último dato real: {last_date.isoformat()} (fila {last_row}).")

    eer = bis_client.get_eer(start_period=(last_date.replace(day=1)).isoformat()[:7])
    codes = [code for _, code in bis_client.COUNTRIES]
    new_dates = sorted(set.intersection(*[set(eer[c]) for c in codes]) - {last_date})
    new_dates = [d for d in new_dates if d > last_date]

    if not new_dates:
        log.append("[TCR] sin novedades (el BIS no tiene un mes más nuevo que el archivo).")
        return False

    any_applied = False
    for d in sorted(new_dates):
        target_row = find_target_row(ws_v, last_row, d)
        if target_row is None:
            log.append(f"[TCR] hay dato nuevo del BIS para {d.isoformat()} pero la hoja no tiene una fila "
                       f"pre-armada con esa fecha; no se agrega automáticamente (requiere revisión).")
            continue
        log.append(f"[TCR] {d.isoformat()}: dato nuevo del BIS encontrado para los 15 países (fila {target_row}).")
        if not apply:
            any_applied = True
            continue
        values_by_col = {col: eer[code][d] for col, (_, code) in zip(COLS, bis_client.COUNTRIES)}
        update_tcr_sheet(xlsx_path, target_row, values_by_col, log)
        chain_values = update_ajustado_sheet(xlsx_path, target_row, log)
        summary = update_ajustado_summary(xlsx_path, target_row, chain_values, log)
        if summary:
            update_ranking_table(xlsx_path, summary, log)
        any_applied = True
        last_row = target_row
        ws_v = openpyxl.load_workbook(xlsx_path, data_only=True)["TCR"]

    return any_applied and apply


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    xlsx_path = Path(args.xlsx)
    pptx_path = Path(args.pptx)
    log = []
    changed = run(xlsx_path, pptx_path, log, args.apply)
    print("\n".join(log))

    if not args.apply:
        print("\n(dry run: no se escribió nada. Corre con --apply para aplicar los cambios.)")
        return
    if not changed:
        print("\nNada que actualizar.")
        return

    touched = pptx_chart_sync.sync_pptx(pptx_path, xlsx_path, pptx_path, verbose=True, no_grow_sheets={"Resumen"})
    print(f"\npptx actualizado ({len(touched)} gráfico(s)): {pptx_path}")
    touched_xlsx = pptx_chart_sync.sync_xlsx_own_charts(xlsx_path, xlsx_path, verbose=True, no_grow_sheets={"Resumen"})
    print(f"Excel (gráficos nativos) actualizado ({len(touched_xlsx)}): {xlsx_path}")


if __name__ == "__main__":
    main()
