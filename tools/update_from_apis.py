#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Trae datos nuevos directo desde las APIs del Banco Central y del INE, y
agrega una fila nueva a las hojas correspondientes de
Graficos 11112026.xlsx cuando corresponde — conservando las fórmulas que ya
tiene el archivo (sólo escribe los valores de entrada; las columnas
calculadas se agregan como fórmula, con su mismo patrón, no como número
fijo). Después sincroniza los gráficos del .pptx y del propio Excel con
pptx_chart_sync.py.

Cómo escribe: nunca abre el libro completo con openpyxl en modo escritura.
Varias columnas "de apoyo" en estas hojas usan fórmulas compartidas de
Excel que openpyxl no reevalúa al guardar — un guardado así deja en blanco
el valor cacheado de esas fórmulas en CUALQUIER hoja del archivo, se haya
tocado o no (confirmado en pruebas). xlsx_row_writer.py evita esto
insertando la fila nueva directamente en el XML de la hoja, sin tocar nada
más del archivo.

Fuentes y alcance (ver tools/README.md):
  - BCCh F051.BT02 / F051.BT03  -> hoja "Balance 2", columnas B y C
  - BCCh F089.PIB.V12.2027.M    -> hoja "PIB 2027", columna C (EEE)
  - INE DF_OCU_SEXO / DF_OCUINF_SEXO -> hoja "Empleo", columnas D y F

Uso:
    python3 tools/update_from_apis.py --xlsx "Graficos 11112026.xlsx" \
        --pptx "Laminas 11112026.pptx" --apply

    Sin --apply, sólo muestra qué encontraría (modo "dry run").
"""
import argparse
import datetime
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
import bcch_client
import ine_client
import xlsx_row_writer as rw
import pptx_chart_sync

MESES_ABR = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def last_data_row(ws, col_letter, start_row):
    r = start_row
    last = None
    while ws[f"{col_letter}{r}"].value is not None:
        last = r
        r += 1
    return last


def plan_balance2(xlsx_path, log):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Balance 2"]
    last_row = last_data_row(ws, "A", 2)
    last_date = ws[f"A{last_row}"].value.date()
    log.append(f"[Balance 2] última fila existente: {last_date.isoformat()} (fila {last_row})")

    ingresos = dict(bcch_client.get_series(bcch_client.SERIES["ingresos_gobierno_central"],
                                            firstdate=(last_date - datetime.timedelta(days=5)).isoformat()))
    gastos = dict(bcch_client.get_series(bcch_client.SERIES["gasto_gobierno_central"],
                                          firstdate=(last_date - datetime.timedelta(days=5)).isoformat()))
    new_dates = sorted(d for d in ingresos if d > last_date and d in gastos)
    if not new_dates:
        log.append("[Balance 2] sin novedades (BCCh no tiene datos más nuevos que el archivo).")
        return []

    # La columna E (PIB nominal) no es un dato plano: en varias filas ya
    # existentes es ella misma una fórmula que proyecta el trimestre
    # (ej. "+E163*(1+3%)", un supuesto de crecimiento t/t) hasta que se
    # reemplaza a mano por el PIB nominal real publicado. No tenemos fuente
    # pública verificada para ese supuesto, así que no lo reproducimos: cada
    # fila nueva llega sin E/G/N/O/P, y chart4/chart5 (que dependen de N/O/P)
    # no se extienden hasta que alguien complete esa columna.
    plans = []
    r = last_row
    for d in new_dates:
        nr = r + 1
        ing, gas = ingresos[d], gastos[d]
        cells = [
            ("A", "date", d, None),
            ("B", "num", ing, None),
            ("C", "num", gas, None),
            ("D", "formula", (f"+B{nr}-C{nr}", ing - gas), None),
            ("M", "date", d, None),
        ]
        plans.append(("Balance 2", nr, cells))
        log.append(f"[Balance 2] fila {nr} a agregar: {d.isoformat()} ingresos={ing:.1f} gasto={gas:.1f} "
                    f"(columnas E/G/N/O/P -PIB nominal y %PIB- quedan pendientes: requieren el supuesto/"
                    f"dato de PIB nominal, sin fuente pública confirmada; chart4/chart5 no se extienden "
                    f"hasta completarlas).")
        r = nr
    return plans


def plan_pib2027(xlsx_path, log):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    wbf = openpyxl.load_workbook(xlsx_path, data_only=False)
    ws, wsf = wb["PIB 2027"], wbf["PIB 2027"]
    last_row = last_data_row(ws, "B", 4)
    last_date = ws[f"B{last_row}"].value.date()
    log.append(f"[PIB 2027] última fila existente: {last_date.isoformat()} (fila {last_row})")

    eee = dict(bcch_client.get_series(bcch_client.SERIES["eee_pib_2027"],
                                       firstdate=(last_date - datetime.timedelta(days=5)).isoformat()))
    new_dates = sorted(d for d in eee if d > last_date)
    if not new_dates:
        log.append("[PIB 2027] sin novedades (la EEE de BCCh no tiene meses más nuevos que el archivo).")
        return []

    plans = []
    r = last_row
    prev_f = ws[f"F{r}"].value
    for d in new_dates:
        nr = r + 1
        cells = [
            ("A", "num", float(d.month), None),
            ("B", "date", d, None),
            ("C", "num", eee[d], None),
        ]
        if prev_f is not None:
            cells.append(("F", "formula", (f"+F{nr-1}", prev_f), None))
        plans.append(("PIB 2027", nr, cells))
        log.append(f"[PIB 2027] fila {nr} a agregar: {d.isoformat()} EEE={eee[d]}. "
                    f"Bloomberg/FMI quedan sin dato (sin fuente pública). Revisa si esta fecha debería "
                    f"ser una barra nueva en el gráfico (requiere editar chart6.xml a mano).")
        r = nr
    return plans


def plan_empleo(xlsx_path, log):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Empleo"]
    last_row = last_data_row(ws, "U", 20)
    last_date = ws[f"U{last_row}"].value.date()
    log.append(f"[Empleo] última fila existente: {last_date.isoformat()} (fila {last_row})")

    data = ine_client.get_formal_informal_total(lookback_quarters=6)
    new_dates = sorted(d for d in data if d > last_date)
    if not new_dates:
        log.append("[Empleo] sin novedades (el INE no tiene trimestres móviles más nuevos que el archivo).")
        return []

    plans = []
    r = last_row
    d_hist = {rr: ws[f"D{rr}"].value for rr in range(max(20, r - 12), r + 1)}
    f_hist = {rr: ws[f"F{rr}"].value for rr in range(max(20, r - 12), r + 1)}
    for d in new_dates:
        nr = r + 1
        start_month = d.month - 2 if d.month > 2 else d.month + 10
        label = f"{MESES_ABR[start_month-1]} - {MESES_ABR[d.month-1]}"
        formal = data[d]["formal_miles"]
        informal = data[d]["informal_miles"]
        d_hist[nr], f_hist[nr] = formal, informal
        cells = [
            ("A", "num", float(d.year), None),
            ("B", "str", label, None),
            ("D", "num", formal, None),
            ("F", "num", informal, None),
            ("U", "date", d, None),
        ]
        if (nr - 12) in d_hist and d_hist[nr - 12] is not None:
            cells.append(("V", "formula", (f"+D{nr}-D{nr-12}", formal - d_hist[nr - 12]), None))
            cells.append(("W", "formula", (f"+F{nr}-F{nr-12}", informal - f_hist[nr - 12]), None))
            cells.append(("X", "formula", (f"+V{nr}+W{nr}", (formal - d_hist[nr - 12]) + (informal - f_hist[nr - 12])), None))
        plans.append(("Empleo", nr, cells))
        log.append(f"[Empleo] fila {nr} a agregar: {d.isoformat()} ({label}) "
                    f"formal={formal:.1f} miles, informal={informal:.1f} miles")
        r = nr
    return plans


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--apply", action="store_true", help="Si no se indica, sólo muestra qué haría (dry run).")
    args = ap.parse_args()

    xlsx_path = Path(args.xlsx)
    pptx_path = Path(args.pptx)

    log = []
    all_plans = []
    for fn in (plan_balance2, plan_pib2027, plan_empleo):
        try:
            all_plans.extend(fn(xlsx_path, log))
        except Exception as e:
            log.append(f"[{fn.__name__}] ERROR: {e}")

    print("\n".join(log))

    if not all_plans:
        print("\nNada que actualizar: todas las fuentes ya están al día con el archivo.")
        return

    if not args.apply:
        print(f"\n(dry run: {len(all_plans)} fila(s) nuevas detectadas, no se escribió nada. "
              f"Corre de nuevo con --apply para escribirlas.)")
        return

    for sheet, row_num, cells in all_plans:
        rw.append_row(xlsx_path, sheet, row_num, cells)
    print(f"\nExcel actualizado ({len(all_plans)} fila(s) nuevas): {xlsx_path}")

    touched = pptx_chart_sync.sync_pptx(pptx_path, xlsx_path, pptx_path, verbose=True)
    print(f"\npptx actualizado ({len(touched)} gráfico(s)): {pptx_path}")
    touched_xlsx = pptx_chart_sync.sync_xlsx_own_charts(xlsx_path, xlsx_path, verbose=True)
    print(f"Excel (gráficos nativos) actualizado ({len(touched_xlsx)}): {xlsx_path}")


if __name__ == "__main__":
    main()
