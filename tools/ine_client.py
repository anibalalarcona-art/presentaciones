"""
Cliente mínimo para la API SDMX del INE (SIMEL / Data Explorer).
No requiere credenciales.

Endpoint: https://sdmx.ine.gob.cl/rest/data/CL01,<dataflow>,1.0/all?format=csv
"""
import csv
import datetime
import io
import urllib.request

BASE_URL = "https://sdmx.ine.gob.cl/rest/data/CL01,{dataflow},1.0/all?format=csv"

# Dataflows confirmados (los totales nacionales calzan exacto con
# Graficos 11112026.xlsx, hoja Empleo, columnas D y F):
DF_OCUPADOS_TOTAL = "DF_OCU_SEXO"       # Población ocupada nacional, por sexo
DF_OCUPADOS_INFORMAL = "DF_OCUINF_SEXO"  # Población ocupada informal nacional, por sexo


def _period_start_date(time_period):
    # formato "2026-07-01/P3M" -> date(2026,7,1)
    date_part = time_period.split("/")[0]
    return datetime.date.fromisoformat(date_part)


def get_national_series(dataflow_id, area_ref="_T", sexo="AS"):
    """
    Devuelve {date: valor} para el total nacional (AREA_REF=_T) y ambos
    sexos (SEXO=AS) de un dataflow del INE. `date` es el primer mes del
    trimestre móvil (p.ej. 2026-07-01 para el trimestre jul-ago-sep).
    """
    url = BASE_URL.format(dataflow=dataflow_id)
    with urllib.request.urlopen(url, timeout=60) as resp:
        raw = resp.read().decode("utf-8")
    reader = csv.DictReader(io.StringIO(raw))
    out = {}
    for row in reader:
        if row.get("AREA_REF") == area_ref and row.get("SEXO") == sexo:
            d = _period_start_date(row["TIME_PERIOD"])
            out[d] = float(row["OBS_VALUE"])
    return out


def get_formal_informal_total(lookback_quarters=6):
    """
    Devuelve, para los últimos `lookback_quarters` trimestres móviles
    disponibles, un dict:
      {quarter_end_date: {"formal_miles": x, "informal_miles": y}}
    donde quarter_end_date es el ÚLTIMO mes del trimestre (la convención
    que usa Graficos 11112026.xlsx, p.ej. 2026-09-01 para "Jul - Sep").
    formal/informal están en MILES de personas (nivel, no variación — la
    variación a/a se calcula en el propio Excel con una fórmula sobre dos
    filas, 12 meses atrás).
    """
    total = get_national_series(DF_OCUPADOS_TOTAL)
    informal = get_national_series(DF_OCUPADOS_INFORMAL)
    common = sorted(set(total) & set(informal))[-lookback_quarters:]
    out = {}
    for start in common:
        end = datetime.date(start.year, start.month, 1)
        # fin del trimestre móvil = inicio + 2 meses
        m = end.month + 2
        y = end.year + (m - 1) // 12
        m = (m - 1) % 12 + 1
        end = datetime.date(y, m, 1)
        informal_miles = informal[start] / 1000.0
        formal_miles = (total[start] - informal[start]) / 1000.0
        out[end] = {"formal_miles": formal_miles, "informal_miles": informal_miles}
    return out
