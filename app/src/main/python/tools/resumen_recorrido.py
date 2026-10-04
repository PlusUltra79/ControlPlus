# -*- coding: utf-8 -*-
"""
resumen_recorrido.py  --  Herramienta 10: Resumen de recorrido / paradas de una UT
=================================================================================
Lee el export de historial de UNA unidad del GTRMax (CSV separado por ';' o
Excel) y genera hasta dos mensajes:

  1) Resumen del recorrido de la UT
  2) Resumen de paradas de la UT

Columnas que se usan del archivo (detectadas por nombre, ignorando tildes y
mayusculas -- el export real trae encabezados en blanco/duplicados y a veces
la codificacion rota, ej. "DirecciÃ³n"):
  - Fecha Reporte : fecha y hora de cada registro (de aqui sale la fecha y las horas)
  - Direccion     : la direccion REAL es la columna que esta justo antes de
                    "Latitud" (la columna llamada "Direccion" solo trae "Venezuela"
                    porque el export parte la direccion en 5 columnas).
  - Velocidad, Tiempo parada (HH:MM:SS, se usa tal cual lo reporta el GTRMax).

Reglas acordadas con el usuario:
  * "Tiempo parada" se usa TAL CUAL; solo cuentan paradas de MAS de 5 minutos.
  * Hora de una parada = hora de la columna "Fecha Reporte" de esa fila.
  * Inicio de operacion = primera fila (en orden cronologico) con velocidad > 0.
    Las paradas anteriores al inicio (la noche en el patio) no se listan.
  * Resguardo: si la unidad tiene resguardo propio (rutas_db.json ->
    PUNTOS_RESGUARDO) se muestra ese; si no, y arranca en un Patio, el nombre
    del patio; si no, la direccion que aparece en el archivo.
  * "Observacion" queda en "-".
  * Saludo segun el reloj del sistema; fecha tomada del archivo.
"""

import io
import os
import re
import json
import math
import logging
import unicodedata
from datetime import datetime, time as _dtime, timedelta

import pandas as pd

from tools import excesos_pdf as _m_excesos
from tools import ojo_de_halcon as _m_ojo

log = logging.getLogger("controlplus_bot")

MIN_PARADA_SEG = 5 * 60          # solo paradas de MAS de 5 minutos
SALTO_IMPOSIBLE_KMH = 150        # descarta registros iniciales con salto fisicamente imposible

# ── Unidades: eje, sub-eje y VIN por defecto (documento de ejes del usuario) ──
_UNIDADES_TXT = """
AM|Norte: R-022 4919, R-045 3916, R-046 4625, R-048 4791, R-053 8126, 0598 0415
AM|Sur: R-047 9924, R-066 2671, R-069 8396, 0349 0589, 0491 8644, 0613 0280, 0615 2295, 0647 0693, 0654 0722, 0664 0745
AM|Centro: R-060 9867, 0219 8792, 0489 8627, 0566 6324
BLV|Acevedo: R-010 5944, 0352 0624, 0452 4436, 0599 0396, 0643 0705, 0645 0684, 0656 0723, 0665 0744, 0680 0179, 0721 0101, 0722 0052
BLV|Brión: R-006 9616, R-007 2273, R-009 7225, R-052 8252, R-058 2109, 0426 4361, 0522 0464, 0641 0703, 0657 0712, 0666 0713, 0700 0031
BLV|Buroz: R-005 9299, R-019 5738, 0416 0132, 0579 0426, 0616 2244
BLV|Páez: R-049 4622, R-050 4874, 0124 0947, 0343 0573, 0481 0781, 0482 4276, 0582 3270, 0609 6312, 0629 8060, 0636 0682, 0642 0691, 0646 0690, 0648 0692
BLV|Pedro Gual: R-008 8319, R-034 6159, R-056 0039, 0477 4434, 0667 0725
MET|Met: R-012 4649, R-013 4955, R-014 4742, R-031 4859, R-064 6656, 0429 4326, 0495 8661, 0497 9519, 0500 9536, 0512 8436, 0516 7824, 0531 0404, 0546 6384, 0581 0423, 0597 0428, 0685 0257
MET|Sucre: R-025 4954, R-033 4665, R-051 3750, 0436 4458, 0511 9921, 0612 6362, 0632 7938
OCM|Tuy I: R-039 4778, R-040 0166, R-041 4765, R-062 1109, 0140 1182, 0424 4343, 0462 4412, 0463 4366, 0485 4413, 0624 7851, 0625 7850, 0705 0108
OCM|Tuy II: R-042 7819, R-043 4769, R-044 2905, R-054 3606, R-055 3949, R-057 7799, R-063 4704, 0417 4470, 0445 4273, 0450 4270, 0454 4505, 0614 2277
PZ|Plaza Zamora: R-001 4918, R-016 4728, R-017 2120, R-018 1419, R-020 8092, R-021 3679, R-023 4118, R-024 3419, R-026 4923, R-027 4718, R-029 4043, R-061 1481, R-067 8148, 0425 4352, 0430 4268, 0448 4370, 0480 4414, 0484 4344, 0503 7815, 0583 0383, 0608 6346, 0637 0696, 0638 5153
SERV|Cont: 0446 4473, 0466 4489, 0592 2186, 0673 6627, 0690 0276, 0697 0275
SERV|Esp: R-011 4650, R-036 4648, 0505 7852, 0514 9545, 0533 6307, 0538 0393, 0557 0385, 0562 6353, 0591 0460, 0594 0452, 0671 6481, 0672 6490, 0702 0016, 0712 0041
"""

_EJE_NOMBRE = {"AM": "Am", "BLV": "Blv", "MET": "Met", "OCM": "Ocm", "PZ": "Pz", "SERV": "Serv"}


def _cargar_unidades():
    db = {}
    for linea in _UNIDADES_TXT.strip().splitlines():
        eje, resto = linea.strip().split("|", 1)
        sub, items = resto.split(": ", 1)
        for it in items.split(", "):
            alias, vin = it.split()
            db[alias] = (_EJE_NOMBRE[eje], sub, vin)
    return db


UNIDADES_DB = _cargar_unidades()

# ── Patios de los Ejes con nombre (documento de patios del usuario) ──────────
PATIOS_NOMBRADOS = [
    ("Patio Vencedores", "Calle Los Teques - San Pedro con Local-7, Avenida Víctor Baptista. Urbanización Santo Omero. Los Teques. Parroquia Los Teques. Municipio Guaicaipuro. Estado Miranda. Venezuela."),
    ("Patio Vencedores", "Local-7, Avenida Víctor Baptista con Calle Los Teques - San Pedro. Urbanización Santo Omero. Los Teques. Parroquia Los Teques. Municipio Guaicaipuro. Estado Miranda. Venezuela."),
    ("Patio Vencedores", "Local-7, Avenida Víctor Baptista con Vía Cárcel de Mujeres. Ramo Verde. Los Teques. Parroquia Los Teques. Municipio Guaicaipuro. Estado Miranda. Venezuela."),
    ("Proveeduría", "Local-7, Avenida Bicentenaria con Calle Ali Primera. Urbanización El Berbech. Los Teques. Parroquia Los Teques. Municipio Guaicaipuro. Estado Miranda. Venezuela."),
    ("Patio Cagua", "Troncal-9, Carretera Caucagua - El Clavo con Calle El Castaño. Panaquire. Caucagua. Parroquia Marizapa. Municipio Acevedo. Estado Miranda. Venezuela."),
    ("Patio Brión", "Calle El Calvario entre Vía El Cien y Avenida Manzanares. Maturín. Tacarigua de Mamporal. Parroquia Tacarigua. Municipio Brión. Estado Miranda. Venezuela."),
    ("Patio Higuerote", "Carretera Higuerote - Curiepe con Troncal-12, Carretera Higuerote - Caucagua. Ciudad Balneario Higuerote. Higuerote. Parroquia Higuerote. Municipio Brión. Estado Miranda. Venezuela."),
    ("Patio Mamporal", "Carretera Tacarigua - Río Chico entre Calle Rincón Bonito y Local-8, Carretera Mamporal - Barlovento. Santo Domingo. Mamporal. Parroquia Mamporal. Municipio Buroz. Estado Miranda. Venezuela."),
    ("Patio Andrés Bello", "Calle Local 8 con Local-8, Avenida Rafael Arevalo Gonzalez. El Delirio. San José de Río Chico. Parroquia San José de Barlovento. Municipio Andrés Bello. Estado Miranda. Venezuela."),
    ("Patio Urbina", "Troncal-9, Autopista Cacique Guaicaipuro entre Avenida Principal de La Urbina y Distribuidor La Urbina. La Urbina. Caracas. Parroquia Petare. Municipio Sucre. Estado Miranda. Venezuela."),
    ("Patio Urbina", "Calle 3 entre Calle 2 - 3 y Avenida Principal de La Urbina. La Urbina. Caracas. Parroquia Petare. Municipio Sucre. Estado Miranda. Venezuela."),
    ("Patio Charallave Sur", "Avenida Gumeral entre Calle 7 de Abril y Calle Manga de Coleo. Barrio Campo Elías. Charallave. Parroquia Charallave. Municipio Cristóbal Rojas. Estado Miranda. Venezuela."),
    ("Patio Ingenio", "Transversal 1 entre Calle 2 y Transversal 3. Sector San José. Guatire. Parroquia Guatire. Municipio Zamora. Estado Miranda. Venezuela."),
    ("Patio Las Rosas", "Calle Principal La Rosa con Calle C. Sector Las Rosas. Guatire. Parroquia Guatire. Municipio Zamora. Estado Miranda. Venezuela."),
]


def _norm_dir(texto):
    """Direccion recortada (sin Parroquia/Municipio/...) y normalizada para comparar."""
    return _m_ojo._normalizar_comparacion(_m_ojo.limpiar_ubicacion(texto))


_PATIOS_NORM = {}
for _nombre, _dir in PATIOS_NOMBRADOS:
    _PATIOS_NORM[_norm_dir(_dir)] = _nombre


def nombre_patio(direccion):
    """Nombre del patio si la direccion coincide con uno; si no, None."""
    return _PATIOS_NORM.get(_norm_dir(direccion))


# ── Unidad ───────────────────────────────────────────────────────────────────

def normalizar_unidad(texto):
    """'r066', 'R-66', 'UT: 343' -> 'R-066', 'R-066', '0343'. None si no es valido."""
    if texto is None:
        return None
    t = _m_ojo.normalizar_alias(texto).upper().replace(" ", "")
    m = re.fullmatch(r"R-?(\d{1,3})", t)
    if m:
        return "R-%03d" % int(m.group(1))
    m = re.fullmatch(r"\d{1,4}", t)
    if m:
        return t.zfill(4)
    return None


def alias_desde_nombre_archivo(nombre):
    """'Unidad_R-053_-_Desde_...csv' -> 'R-053' (o None)."""
    m = re.search(r"(?i)unidad[_\s-]+(R-?\d{1,3}|\d{3,4})", nombre or "")
    return normalizar_unidad(m.group(1)) if m else None


def buscar_unidad(alias):
    """(eje, sub_eje, vin) de la unidad, o None si no esta en la lista."""
    return UNIDADES_DB.get(alias)


def parsear_unidad_manual(texto):
    """Para unidades que no estan en la lista: 'Blv, Acevedo, 1234' -> (eje, sub, vin) o None."""
    partes = [p.strip() for p in re.split(r"[,;]", texto or "") if p.strip()]
    if len(partes) == 2 and partes[0].lower() == "pz":
        partes = [partes[0], "Plaza Zamora", partes[1]]
    if len(partes) != 3:
        return None
    eje = _EJE_NOMBRE.get(partes[0].upper())
    if not eje:
        return None
    sub = partes[1]
    vin = partes[2].upper()
    if not re.fullmatch(r"\d{4}|S/N", vin):
        return None
    # usar el nombre oficial del sub-eje si existe (ej. 'tuy ii' -> 'Tuy II')
    oficiales = {v[1].lower(): v[1] for v in UNIDADES_DB.values() if v[0] == eje}
    sub = oficiales.get(sub.lower(), sub.title())
    return (eje, sub, vin)


def _linea_eje(eje, sub):
    if eje == "Pz":
        return "- Eje *Pz*"
    if eje == "Serv":
        return "- Eje *Serv. Cont*" if sub.strip().lower() == "cont" else "- Eje *Serv Esp*"
    if eje == "Met":
        sub = "Met"   # mismo criterio que Data Alertas: Sucre -> Met
    return f"- Eje *{eje}* - Sub eje *{sub}*"


# ── Lectura del archivo ──────────────────────────────────────────────────────

def _norm_h(h):
    s = unicodedata.normalize("NFKD", str(h))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _leer_tabla(ruta):
    ext = os.path.splitext(ruta)[1].lower()
    if ext in (".csv", ".txt"):
        contenido = None
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                with open(ruta, encoding=enc) as f:
                    contenido = f.read()
                break
            except UnicodeDecodeError:
                continue
        primera = contenido.split("\n", 1)[0]
        sep = ";" if primera.count(";") >= primera.count(",") else ","
        return pd.read_csv(io.StringIO(contenido), sep=sep, dtype=str, keep_default_na=False)
    try:
        return pd.read_excel(ruta, engine="calamine", dtype=object)
    except Exception:
        return pd.read_excel(ruta, dtype=object)


def _vacio(v):
    return v is None or (isinstance(v, float) and math.isnan(v)) or (isinstance(v, str) and not v.strip())


def _parse_dt(v):
    if _vacio(v):
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return pd.Timestamp(v).to_pydatetime()
    s = str(v).replace("\u00a0", " ").replace("\u202f", " ").strip().lower()
    m = re.fullmatch(
        r"(\d{1,2})/(\d{1,2})/(\d{4})[ ,]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(?:([ap])\.?\s*m\.?)?", s)
    if not m:
        return None
    d, mo, y, h, mi, se, ap = m.groups()
    h = int(h)
    if ap:
        h = h % 12 + (12 if ap == "p" else 0)
    return datetime(int(y), int(mo), int(d), h, int(mi), int(se or 0))


def _parse_duracion(v):
    """'00:09:45' / 'MM:SS' / time / timedelta / fraccion de dia -> segundos (o None)."""
    if _vacio(v):
        return None
    if isinstance(v, timedelta):
        return int(v.total_seconds())
    if isinstance(v, _dtime):
        return v.hour * 3600 + v.minute * 60 + v.second
    if isinstance(v, (int, float)):
        return int(round(float(v) * 86400)) if float(v) < 1 else int(v)
    m = re.fullmatch(r"(?:(\d+):)?(\d{1,2}):(\d{2})", str(v).strip())
    if not m:
        return None
    h, mi, se = m.groups()
    return int(h or 0) * 3600 + int(mi) * 60 + int(se)


def _parse_num(v):
    if _vacio(v):
        return None
    m = re.search(r"-?\d+(?:[.,]\d+)?", str(v))
    return float(m.group(0).replace(",", ".")) if m else None


def _haversine_km(lat1, lon1, lat2, lon2):
    r = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = p2 - p1, math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def leer_recorrido(ruta, log_fn=print):
    """
    Devuelve un DataFrame ordenado cronologicamente con columnas:
    dt, vel, tp_seg, dir (completa), dir_limpia, lat, lon
    """
    df = _leer_tabla(ruta)
    cols = list(df.columns)
    ncols = [_norm_h(c) for c in cols]

    def _idx(nombre):
        return ncols.index(nombre) if nombre in ncols else None

    i_fecha, i_tp = _idx("fechareporte"), _idx("tiempoparada")
    i_vel, i_lat, i_lon = _idx("velocidad"), _idx("latitud"), _idx("longitud")
    faltan = [n for n, i in (("Fecha Reporte", i_fecha), ("Tiempo parada", i_tp),
                             ("Velocidad", i_vel), ("Latitud", i_lat)) if i is None]
    if faltan:
        raise ValueError(f"Columnas no encontradas: {', '.join(faltan)}. "
                         f"Columnas detectadas: {[str(c).strip() for c in cols]}")

    i_dir = i_lat - 1
    # la direccion real es la ultima columna de texto antes de Latitud; verificar
    def _largo(i):
        return df.iloc[:, i].astype(str).str.len().mean()
    if _largo(i_dir) < 15:
        candidatas = range(0, i_lat)
        i_dir = max(candidatas, key=_largo)
    log_fn(f"  Columna de direccion usada: posicion {i_dir + 1} (encabezado '{str(cols[i_dir]).strip()}')")

    filas = []
    for _, r in df.iterrows():
        dt = _parse_dt(r.iloc[i_fecha])
        if dt is None:
            continue
        d = "" if _vacio(r.iloc[i_dir]) else str(r.iloc[i_dir]).strip()
        filas.append({
            "dt": dt,
            "vel": _parse_num(r.iloc[i_vel]) or 0.0,
            "tp_seg": _parse_duracion(r.iloc[i_tp]),
            "dir": d,
            "dir_limpia": _m_ojo.limpiar_ubicacion(d),
            "lat": _parse_num(r.iloc[i_lat]),
            "lon": _parse_num(r.iloc[i_lon]) if i_lon is not None else None,
        })
    if not filas:
        raise ValueError("No se pudo leer ninguna fila con 'Fecha Reporte' valida.")

    # El export viene del mas reciente al mas antiguo: invertir para que los
    # empates de segundo conserven el orden real, y luego ordenar (estable).
    if filas[0]["dt"] > filas[-1]["dt"]:
        filas.reverse()
    filas.sort(key=lambda f: f["dt"])

    # Registros iniciales con salto imposible (ej. 12 km en 3 min): se descartan
    descartadas = 0
    while len(filas) > 1:
        a, b = filas[0], filas[1]
        if None in (a["lat"], a["lon"], b["lat"], b["lon"]):
            break
        horas = max((b["dt"] - a["dt"]).total_seconds(), 1) / 3600
        if _haversine_km(a["lat"], a["lon"], b["lat"], b["lon"]) / horas > SALTO_IMPOSIBLE_KMH:
            filas.pop(0)
            descartadas += 1
        else:
            break
    if descartadas:
        log_fn(f"  ⚠ {descartadas} registro(s) inicial(es) descartado(s) por salto de posicion imposible.")

    out = pd.DataFrame(filas)
    fechas = sorted({d.date() for d in out["dt"]})
    if len(fechas) > 1:
        log_fn(f"  ⚠ El archivo abarca {len(fechas)} dias; se usa la fecha del primer registro.")
    log_fn(f"  Registros leidos: {len(out)} ({out['dt'].iloc[0]:%d/%m/%Y %H:%M} a {out['dt'].iloc[-1]:%H:%M})")
    return out


# ── Analisis ─────────────────────────────────────────────────────────────────

def _hora12(dt):
    h = dt.hour % 12 or 12
    return f"{h:02d}:{dt.minute:02d} {'AM' if dt.hour < 12 else 'PM'}"


def _duracion(seg):
    return f"{seg // 60}:{seg % 60:02d} minutos"


def _resguardo_propio(alias, ruta_rutas_json):
    """Direcciones (limpias) de resguardo propias de la unidad, segun rutas_db.json."""
    try:
        with open(ruta_rutas_json, encoding="utf-8") as f:
            db = json.load(f)
    except Exception:
        return []
    dirs = (db.get("PUNTOS_RESGUARDO") or {}).get(str(alias).upper(), [])
    return [_m_ojo.limpiar_ubicacion(d) for d in dirs]


def _lugar(direccion_limpia):
    """Nombre del patio si aplica; si no, la direccion."""
    return nombre_patio(direccion_limpia) or direccion_limpia


def _clave_lugar(direccion_limpia):
    return nombre_patio(direccion_limpia) or _m_ojo._normalizar_comparacion(direccion_limpia)


def analizar(df, alias, ruta_rutas_json):
    """Extrae del historial todo lo que necesitan los dos reportes."""
    fecha = df["dt"].iloc[0]
    movs = df.index[df["vel"] > 0].tolist()
    res = {"alias": alias, "fecha": fecha, "inicio": None}

    if movs:
        i0 = movs[0]
        res["inicio"] = df["dt"].iloc[i0]
        res["dir_inicio"] = _lugar(df["dir_limpia"].iloc[i0])
        previo = df.iloc[max(i0 - 1, 0)]
        dir_previa = previo["dir_limpia"]
        propios = _resguardo_propio(alias, ruta_rutas_json)
        if propios:
            res["resguardo"] = " / ".join(dict.fromkeys(propios))
        else:
            res["resguardo"] = _lugar(dir_previa)
        paradas = df.iloc[i0:]
    else:
        propios = _resguardo_propio(alias, ruta_rutas_json)
        res["resguardo"] = " / ".join(dict.fromkeys(propios)) if propios else _lugar(df["dir_limpia"].iloc[0])
        paradas = df.iloc[0:0]

    res["paradas"] = [
        {"dt": r.dt, "seg": int(r.tp_seg), "dir": r.dir_limpia}
        for r in paradas.itertuples()
        if r.tp_seg is not None and r.tp_seg > MIN_PARADA_SEG
    ]

    # ubicacion actual y desde cuando
    ult = df.iloc[-1]
    clave = _clave_lugar(ult["dir_limpia"])
    j = len(df) - 1
    while j > 0 and _clave_lugar(df["dir_limpia"].iloc[j - 1]) == clave:
        j -= 1
    res["actual"] = _lugar(ult["dir_limpia"])
    res["actual_desde"] = df["dt"].iloc[j]
    return res


# ── Mensajes ─────────────────────────────────────────────────────────────────

def _encabezado(saludo, unidad_info, alias, cts):
    eje, sub, vin = unidad_info
    return [f"> {saludo}", _linea_eje(eje, sub),
            f"- Unidad *{alias}* - Vin *{vin}*", f"- Cts: *{cts}*"]


def _bloque_paradas(paradas):
    bloques = []
    for p in paradas:
        bloques.append(f"*{_hora12(p['dt'])}* - duración de parada *{_duracion(p['seg'])}*\n{p['dir']}")
    return "\n\n".join(bloques)


def mensaje_recorrido(res, unidad_info, cts, saludo):
    alias = res["alias"]
    lineas = _encabezado(saludo, unidad_info, alias, cts)
    lineas += ["", f"*Fecha:* {res['fecha']:%d/%m/%Y}", ""]
    if res["inicio"] is None:
        lineas += ["*Hora de inicio:* -", "",
                   f"*Acción:* La unidad {alias} no registró movimiento en el período.", ""]
    else:
        lineas += [f"*Hora de inicio:* {_hora12(res['inicio'])}", "",
                   f"*Acción:* La unidad {alias} inició movimiento desde *{res['dir_inicio']}*.", ""]
    lineas += [f"*Ubicación regular de resguardo:* {res['resguardo']}", ""]
    if res["paradas"]:
        lineas += ["*Paradas:*", _bloque_paradas(res["paradas"]), ""]
    else:
        lineas += ["*Paradas:* No realizó ninguna.", ""]
    lineas += [f"*Ubicación actual:* Se encuentra ubicada en {res['actual']}.", "",
               "*Observación:* -"]
    return "\n".join(lineas)


def mensaje_paradas(res, unidad_info, cts, saludo):
    alias = res["alias"]
    lineas = _encabezado(saludo, unidad_info, alias, cts)
    lineas += [""]
    lineas += [f"*Inicio de operación:* {_hora12(res['inicio']) if res['inicio'] else '-'}",
               f"*Lugar de resguardo:* {res['resguardo']}"]
    if res["paradas"]:
        lineas += ["*Paradas:*", _bloque_paradas(res["paradas"])]
    else:
        lineas += ["*Paradas:* No realizó ninguna."]
    lineas += ["", f"*Ubicación actual:* {res['actual']} desde las *{_hora12(res['actual_desde'])}*"]
    return "\n".join(lineas)


# ── Funcion publica usada por el bot ─────────────────────────────────────────

def generar_resumen(ruta, alias, cts, tipo, ruta_rutas_json, unidad_info=None,
                    log_fn=print, ahora=None):
    """
    ruta        : CSV/Excel del historial de la unidad
    alias       : alias ya normalizado (normalizar_unidad)
    cts         : nombre escrito por el usuario (se formatea 'Nombre I.')
    tipo        : 'recorrido' | 'paradas' | 'ambos'
    unidad_info : (eje, sub_eje, vin); si es None se busca en UNIDADES_DB
    ahora       : datetime para el saludo (por defecto, reloj del sistema)
    Retorna list[str] con 1 o 2 mensajes (recorrido primero).
    """
    unidad_info = unidad_info or buscar_unidad(alias)
    if unidad_info is None:
        raise ValueError(f"La unidad {alias} no esta en la lista de ejes/VIN.")
    ahora = ahora or datetime.now()
    saludo = _m_excesos.saludo_por_hora(ahora.hour, ahora.minute)
    cts_fmt = _m_excesos.formatear_cts(cts)

    df = leer_recorrido(ruta, log_fn)
    res = analizar(df, alias, ruta_rutas_json)
    log_fn(f"  Inicio: {res['inicio']} | paradas > 5 min: {len(res['paradas'])} | actual: {res['actual']}")

    mensajes = []
    if tipo in ("recorrido", "ambos"):
        mensajes.append(mensaje_recorrido(res, unidad_info, cts_fmt, saludo))
    if tipo in ("paradas", "ambos"):
        mensajes.append(mensaje_paradas(res, unidad_info, cts_fmt, saludo))
    return mensajes
