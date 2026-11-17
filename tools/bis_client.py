"""
Cliente mínimo para la API SDMX del BIS (Bank for International Settlements),
el mismo portal enlazado en la hoja "TCR" de Gráficos_TC.xlsx. No requiere
credenciales.

Serie usada: WS_EER (Effective Exchange Rates), llave FREQ.EER_TYPE.
EER_BASKET.REF_AREA = M.R.B.{código de país} — mensual, real, canasta
amplia (64 economías) — igual al filtro del enlace que ya tenía la hoja
(FREQ=M, EER_BASKET=B, EER_TYPE=R).
"""
import csv
import datetime
import io
import urllib.request

BASE_URL = "https://stats.bis.org/api/v1/data/WS_EER/M.R.B.{areas}/all?startPeriod={start}&format=csv"

# Orden de columnas B:P de la hoja "TCR" (y "TCR (ajustado)"), con su
# código REF_AREA en el BIS. Confirmado: los valores que devuelve la API
# calzan exacto con los que ya había en el Excel para cada país.
COUNTRIES = [
    ("Chile", "CL"),
    ("Brasil", "BR"),
    ("Colombia", "CO"),
    ("México", "MX"),
    ("Argentina", "AR"),
    ("Europa", "XM"),
    ("Reino Unido", "GB"),
    ("China", "CN"),
    ("Japón", "JP"),
    ("Corea", "KR"),
    ("Canadá", "CA"),
    ("Suecia", "SE"),
    ("Dinamarca", "DK"),
    ("Noruega", "NO"),
    ("EE.UU.", "US"),
]


def get_eer(start_period=None):
    """
    Devuelve {area_code: {date: value}} para los 15 países/áreas de
    COUNTRIES, en una sola consulta. `date` es el último día del mes
    (para calzar con la columna A de la hoja "TCR", que usa fin de mes).
    `start_period` en formato "YYYY-MM"; por defecto trae ~2 años.
    """
    if start_period is None:
        start_period = (datetime.date.today().replace(day=1) - datetime.timedelta(days=730)).strftime("%Y-%m")
    areas = "+".join(code for _, code in COUNTRIES)
    url = BASE_URL.format(areas=areas, start=start_period)
    with urllib.request.urlopen(url, timeout=60) as resp:
        raw = resp.read().decode("utf-8")
    reader = csv.DictReader(io.StringIO(raw))
    out = {code: {} for _, code in COUNTRIES}
    for row in reader:
        area = row["REF_AREA"]
        if area not in out:
            continue
        y, m = row["TIME_PERIOD"].split("-")
        y, m = int(y), int(m)
        next_m = datetime.date(y + (m // 12), (m % 12) + 1, 1)
        month_end = next_m - datetime.timedelta(days=1)
        out[area][month_end] = float(row["OBS_VALUE"])
    return out
