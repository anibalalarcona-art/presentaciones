"""
Cliente mínimo para la API REST del Banco Central de Chile (Base de Datos
Estadísticos / SIETE). Usa autenticación user+pass (las credenciales de tu
cuenta BDE), leídas de las variables de entorno BCCH_USER / BCCH_PASS.

No depende de la librería oficial `bcchapi` (requiere un API Key Token
aparte); esto habla directamente con el mismo endpoint REST vía user/pass,
que es el modo de autenticación ya disponible en este entorno.
"""
import datetime
import json
import os
import urllib.parse
import urllib.request

BASE_URL = "https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx"


class BcchError(RuntimeError):
    pass


def _creds():
    user = os.environ.get("BCCH_USER")
    pw = os.environ.get("BCCH_PASS")
    if not user or not pw:
        raise BcchError("Faltan las variables de entorno BCCH_USER / BCCH_PASS")
    return user, pw


def _call(params):
    user, pw = _creds()
    q = {"user": user, "pass": pw}
    q.update(params)
    url = BASE_URL + "?" + urllib.parse.urlencode(q)
    with urllib.request.urlopen(url, timeout=30) as resp:
        raw = resp.read().decode("latin-1")
    data = json.loads(raw)
    if data.get("Codigo") not in (0, None):
        raise BcchError(f"BCCh API error {data.get('Codigo')}: {data.get('Descripcion')}")
    return data


def get_series(series_id, firstdate=None, lastdate=None):
    """
    Devuelve [(date, float_value_o_None), ...] ordenado por fecha, para una
    serie de la Base de Datos Estadísticos del Banco Central.
    """
    params = {"function": "GetSeries", "timeseries": series_id}
    if firstdate:
        params["firstdate"] = firstdate
    if lastdate:
        params["lastdate"] = lastdate
    data = _call(params)
    obs = data["Series"]["Obs"] or []
    out = []
    for o in obs:
        d = datetime.datetime.strptime(o["indexDateString"], "%d-%m-%Y").date()
        v = o["value"]
        val = None if v in (None, "NaN", "") else float(v)
        out.append((d, val))
    out.sort(key=lambda t: t[0])
    return out


def latest(series_id, lookback_days=120):
    """Último valor no nulo de una serie, como (date, value)."""
    firstdate = (datetime.date.today() - datetime.timedelta(days=lookback_days)).isoformat()
    obs = [o for o in get_series(series_id, firstdate=firstdate) if o[1] is not None]
    if not obs:
        raise BcchError(f"Sin observaciones recientes para {series_id}")
    return obs[-1]


# Series ya confirmadas (valores calzan exacto con Graficos 11112026.xlsx):
SERIES = {
    "ingresos_gobierno_central": "F051.BT02.FLU.C.D.CLP.M",   # Gobierno central total, ingresos devengado (millones $)
    "gasto_gobierno_central": "F051.BT03.FLU.C.D.CLP.M",      # Gobierno central total, gastos devengado (millones $)
    "eee_pib_2027": "F089.PIB.V12.2027.M",                    # EEE, expectativa de crecimiento del PIB 2027 (%)
    "pib_nominal_trimestral": "F032.PIB.FLU.N.CLP.EP18.Z.Z.0.T",  # PIB nominal (Cuentas Nacionales), miles de millones $, trimestral
    "dolar_observado": "F073.TCO.PRE.Z.D",                    # usado solo como ejemplo/validación
}
