r"""
tools/ojo_de_halcon.py
-----------------------
Herramienta 9 — Ojo de Halcón (validación en tiempo real de UT)

Cruza el 'Reporte de Estatus Actual' del GTRMax (Alias, Matrícula, Ubicación,
Vel., Km del día, y opcionalmente Último Reporte) con el Excel de
'Disponibilidad' (Estatus, CTS, Ruta) y con un registro interno de Rutas y
su recorrido correcto (rutas_db.json), para generar 5 tipos de incidencia:

  - Fuera de ruta           -> ubicación actual no coincide con la ruta
                               asignada, O coincide con una ubicación de
                               Patio de los Ejes (nuevo v3.2)
  - Movimiento no reportado -> Inoperativa + Km del día > 1.0
  - Exceso de velocidad     -> Vel >= 87 km/h
  - Sin movimiento (v3.2)   -> Vel == 0 km/h
  - UT pegada (v3.2)        -> columna 'Último Reporte' > 1 hora (si la
                               columna existe en el export -- se omite el
                               chequeo sin error si no viene en el archivo)

Sigue las mismas convenciones que el resto de tools/*.py:
  - Firmas públicas estables, log_fn para reportar avisos sin romper el flujo.
  - Reutiliza SUBEJE_A_EJE (misma tabla que ultimas_ut_movimiento.py) y el
    mismo criterio de limpieza de ubicación (limpiar_ubicacion) que esa
    herramienta, con el mismo patrón anclado post-corrección v3.0
    (r"(?:^|\.\s*)PALABRA\b").
  - CTS se formatea igual que en excesos_pdf.py: formatear_cts("Nombre I.").
  - VIN se busca en el mismo VIN_DB compartido (sin_informacion.py /
    excesos_pdf.py) -- se importa perezosamente para no duplicar el diccionario.
  - Patio de los Ejes se compara con el mismo criterio tolerante que
    es_ubicacion_patio() de ultimas_ut_movimiento.py (ignora mayúsculas,
    espacios extra y el punto final).
"""

import re
import json
import unicodedata
import difflib
import numbers
from datetime import datetime

import pandas as pd

# Reutiliza lo ya existente en el proyecto en vez de duplicar lógica de
# negocio (VIN_DB compartido, saludo por hora, formato de nombre de CTS) --
# mismo criterio que exige el protocolo del proyecto para no romper nada.
from tools import excesos_pdf as _m_excesos_pdf

VIN_DB = _m_excesos_pdf.VIN_DB
get_vin_compartido = _m_excesos_pdf.get_vin
saludo_por_hora = _m_excesos_pdf.saludo_por_hora
formatear_cts = _m_excesos_pdf.formatear_cts

UMBRAL_ULTIMO_REPORTE_SEG = 3600       # 1 hora -- piso de "UT pegada"
UMBRAL_ULTIMO_REPORTE_MAX_SEG = 12 * 3600  # 12 horas -- techo de "UT pegada" (nuevo)
# más de 12h se omite: se asume que otra gestión ya se encarga de esas unidades

UMBRAL_SIN_MOVIMIENTO_SEG = 30 * 60            # 30 min -- "Validación Parada" (nuevo v3.4)
UMBRAL_SIN_MOVIMIENTO_FUERA_RUTA_SEG = 10 * 60  # 10 min si además está Fuera de ruta
UMBRAL_SIN_MOVIMIENTO_MAX_SEG = 12 * 3600       # 12 horas -- techo (nuevo v3.5)
# más de 12h se omite, mismo criterio que el techo de "UT pegada"


def _formatear_duracion_legible(segundos):
    """segundos -> '1 hora con 01 min.' / '45 min.' -- para el texto de
    Validación Parada ('(1 hora con 01 min. aproximadamente)')."""
    segundos = int(segundos)
    dias, resto = divmod(segundos, 86400)
    horas, resto = divmod(resto, 3600)
    minutos = resto // 60
    partes = []
    if dias:
        partes.append(f"{dias} día{'s' if dias != 1 else ''}")
    if horas:
        partes.append(f"{horas} hora{'s' if horas != 1 else ''}")
    if partes:
        return " ".join(partes) + f" con {minutos:02d} min."
    return f"{minutos} min."

# ---------------------------------------------------------------------------
# Tabla de mapeo etiqueta (Matrícula) -> Eje / hoja de Disponibilidad
# (misma tabla documentada para ultimas_ut_movimiento.py, extendida con
#  el nombre exacto de la hoja de Disponibilidad que corresponde a cada
#  etiqueta real)
# ---------------------------------------------------------------------------
ETIQUETA_INFO = {
    # etiqueta_normalizada: (Eje, SubEje_o_None, hoja_disponibilidad)
    "norte":      ("Am", None, "Am "),
    "sur":        ("Am", None, "Am "),
    "centro":     ("Am", None, "Am "),
    "acevedo":    ("Blv", "Acevedo", "Acevedo"),
    "brion":      ("Blv", "Brión", "Brion"),
    "brión":      ("Blv", "Brión", "Brion"),
    "buroz":      ("Blv", "Buroz", "Buroz"),
    "paez":       ("Blv", "Páez", "Paez"),
    "páez":       ("Blv", "Páez", "Paez"),
    "pedro gual": ("Blv", "Pedro Gual", "Pedro Gual"),
    "met":        ("Met", None, "Met"),
    "suc":        ("Met", None, "Met"),
    "sucre":      ("Met", None, "Met"),
    "tuy i":      ("Ocm", None, "Ocm"),
    "tuy ii":     ("Ocm", None, "Ocm"),
    "pz":         ("Pz", None, "Pz"),
    "cont":       ("Serv Cont", None, "Serv Cont"),
    "ggm":        ("Serv Cont", None, "Serv Cont"),
    "cttd":       ("Serv Cont", None, "Serv Cont"),
    "esp":        ("Serv Esp", None, "Serv Esp"),
    "bus esc":    ("Serv Esp", None, "Serv Esp"),
}

# Hojas de Disponibilidad sin columna de Estatus utilizable (se omiten sin
# interrumpir el proceso, igual que documenta Arquitectura General).
# Nota (corregido 16/09/2026): 'paez' se sacó de este conjunto -- la hoja
# 'Paez' venía documentada como vacía en versiones anteriores del Excel real,
# pero el archivo real actual SÍ trae una tabla normal (UT/Estatus/Ruta/CTS)
# con datos utilizables. Al tenerla en este set, leer_disponibilidad() la
# saltaba por completo y las 7 unidades del sub-eje Páez del Reporte de
# Estatus Actual quedaban siempre "sin registro en Disponibilidad" (por lo
# tanto, o se omitían, o caían en la incidencia "Sin información" si tenían
# Km del día > 1.0, sin importar que sí estuvieran cargadas en Disponibilidad).
# 'buroz' sigue en el set porque esa hoja sí sigue en formato de texto libre
# (bloques "Ut:XXXX"/"Cts: Nombre"), sin columna de Estatus/Ruta tabulada.
HOJAS_SIN_DATOS_UTILES = {"buroz"}

UMBRAL_VELOCIDAD_KMH = 87
UMBRAL_KM_DIA_SIN_REPORTE = 1.0

# Valores de la columna Ruta/Servicio de Disponibilidad que significan
# "sin ruta/servicio asignado" -- no se evalúa ni "Fuera de ruta" ni "Ruta
# no reconocida" (nuevo) para estos casos: no es que la ruta no se
# reconozca, es que no hay ninguna ruta que evaluar.
# "SS" (sin puntos) se agregó como alias de "S.S" -- mismo significado
# (sin servicio), solo que en algunas hojas se escribe sin puntuación.
_RUTA_NO_ASIGNADA = ("S.R", "S-OP", "S/OP", "S.S", "SS", "")

# Palabras clave que indican que el texto de la columna Ruta es en realidad
# una NOTA operativa/administrativa (por qué la unidad no circula, o a qué
# tarea fue asignada) y no el nombre de una ruta -- se usan para NO generar
# "Ruta no reconocida" en esos casos (pedido explícito del usuario: solo
# avisar cuando el texto parece un nombre de ruta real). Se comparan sin
# tildes/mayúsculas contra el texto completo de la columna Ruta.
# Nota: si el texto combina una ruta real + una nota separadas por "/"
# (ej. "El Alto - La Ceiba/ unidad busca repuestos"), cualquiera de estas
# palabras en la nota hace que el texto COMPLETO se trate como nota y se
# omita -- no se intenta separar la parte de ruta real de la nota.
NOTAS_OPERATIVAS_KEYWORDS = (
    "falla", "resguardo", "repuesto", "gerencia", "disposicion",
    "proveeduria", "admon", "administracion", "coordinador", "apoyo",
    "auxilio",
)


def _es_nota_operativa(texto_ruta):
    """True si el texto de Ruta parece una nota operativa/administrativa
    (no el nombre de una ruta real) -- ver NOTAS_OPERATIVAS_KEYWORDS."""
    t = _normalizar_comparacion(texto_ruta)
    return any(kw in t for kw in NOTAS_OPERATIVAS_KEYWORDS)


# ---------------------------------------------------------------------------
# Saludo por hora (se toma SIEMPRE del reloj del sistema -- herramienta en
# tiempo real, a diferencia de Excesos de Velocidad donde el usuario la teclea).
# Reutiliza saludo_por_hora() de excesos_pdf.py (importado arriba).
# ---------------------------------------------------------------------------
def saludo_actual():
    ahora = datetime.now()
    return saludo_por_hora(ahora.hour, ahora.minute)


# ---------------------------------------------------------------------------
# Limpieza de ubicación (mismo criterio anclado que ultimas_ut_movimiento.py
# tras la corrección v3.0 del bug "Barrio Venezuela")
# ---------------------------------------------------------------------------
_PALABRAS_ADMIN = ("Parroquia", "Municipio", "Estado", "Venezuela")


def limpiar_ubicacion(texto):
    """Corta antes del bloque administrativo final (Parroquia/Municipio/
    Estado/Venezuela). Prioriza SIEMPRE un corte anclado después de un
    punto (el cierre real de la dirección) sobre uno anclado al inicio del
    texto -- bug real corregido en v3.2: varias direcciones del registro de
    rutas ahora traen un prefijo "Estado X Municipio Y Parroquia Z" al
    INICIO (no solo al final, como contexto administrativo redundante).
    Si el corte se buscara indistintamente al inicio o después de un punto
    (criterio original), ese prefijo se detectaba primero y la dirección
    completa quedaba vacía. Ahora el corte al inicio del texto solo se usa
    como último recurso, cuando NO aparece ningún cierre administrativo
    anclado en un punto en ningún lugar del texto (caso documentado: la
    ubicación es únicamente el nombre de la parroquia)."""
    if not texto:
        return ""
    texto = str(texto).strip()

    def _buscar(patron_prefijo):
        mejor = None
        for palabra in _PALABRAS_ADMIN:
            patron = re.compile(patron_prefijo + re.escape(palabra) + r"\b")
            m = patron.search(texto)
            if m and (mejor is None or m.start() < mejor):
                mejor = m.start()
        return mejor

    mejor_corte = _buscar(r"\.\s*")  # 1) preferido: cierre real tras un punto
    if mejor_corte is None:
        mejor_corte = _buscar(r"^")  # 2) último recurso: inicio del texto

    if mejor_corte is None:
        return texto.strip().rstrip(".").strip()
    recorte = texto[:mejor_corte].strip().rstrip(".").strip()
    if recorte:
        return recorte
    # Si no queda nada antes de la palabra administrativa, usar esa palabra
    resto = texto[mejor_corte:].lstrip(". ")
    # tomar solo el primer "segmento" (hasta el próximo punto)
    return resto.split(".")[0].strip()


def _normalizar_comparacion(texto):
    """minúsculas, sin acentos, sin espacios repetidos, sin punto final."""
    if not texto:
        return ""
    t = texto.strip().rstrip(".").strip().lower()
    t = unicodedata.normalize("NFKD", t)
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = re.sub(r"\s+", " ", t)
    return t


# ---------------------------------------------------------------------------
# CTS formateado "Nombre I." -- reutiliza formatear_cts() de excesos_pdf.py
# (importado arriba), agregando solo el mapeo de valores vacíos/"Sin Cts" a
# "S.I." que es específico de esta herramienta (excesos_pdf.py recibe el
# nombre ya filtrado por get_cts(), acá se filtra antes de formatear).
# ---------------------------------------------------------------------------
def formatear_cts_ojo(nombre):
    if not nombre:
        return "S.I."
    nombre = str(nombre).strip()
    if not nombre or nombre.lower() in ("sin cts", "***", "nan"):
        return "S.I."
    return formatear_cts(nombre)


# ---------------------------------------------------------------------------
# Extraer Eje/Sub-eje/hoja a partir de la columna Matrícula
# ("Vin: 8092 Pz" / "VIN: 8396 Sur" -> etiqueta "Pz" / "Sur")
# ---------------------------------------------------------------------------
def extraer_info_eje(matricula, log_fn=print):
    if not matricula:
        return None
    texto = str(matricula)
    texto = re.sub(r"(?i)^\s*vin\s*:\s*\d+\s*", "", texto).strip()
    etiqueta_norm = texto.strip().lower()
    info = ETIQUETA_INFO.get(etiqueta_norm)
    if info is None:
        log_fn(f"⚠ etiqueta sin clasificar en Matrícula: '{matricula}'")
        return None
    eje, subeje, hoja = info
    return {"eje": eje, "subeje": subeje, "hoja_disponibilidad": hoja, "etiqueta": texto.strip()}


# ---------------------------------------------------------------------------
# Lectura del Reporte de Estatus Actual (HTML disfrazado de .xls -> lxml,
# mismo motor que usa ultimas_ut_movimiento.py y por la misma razón: filas
# <tr> sin cerrar que rompen html.parser)
# ---------------------------------------------------------------------------
def leer_estatus_actual(ruta_archivo, log_fn=print):
    tablas = pd.read_html(ruta_archivo, flavor="lxml")
    tabla_datos = max(tablas, key=lambda t: t.shape[0] * t.shape[1])
    tabla_datos.columns = [str(c).strip() for c in tabla_datos.iloc[0]]
    df = tabla_datos.iloc[1:].reset_index(drop=True)

    def col(*nombres):
        for n in nombres:
            if n in df.columns:
                return n
        return None

    c_alias = col("Alias")
    c_matricula = col("Matrícula", "Matricula")
    c_ubicacion = col("Ubicación", "Ubicacion")
    c_vel = col("Vel.", "Vel", "Velocidad")
    c_km = col("Km del día", "Km del dia")
    c_ultimo_reporte = col("Último Reporte", "Ultimo Reporte")

    faltantes = [n for n, c in [("Alias", c_alias), ("Matrícula", c_matricula),
                                 ("Ubicación", c_ubicacion), ("Vel.", c_vel),
                                 ("Km del día", c_km)] if c is None]
    if faltantes:
        raise ValueError(f"Columnas no encontradas en el reporte: {faltantes}")

    out = pd.DataFrame({
        "Alias": df[c_alias].astype(str).str.strip(),
        "Matricula": df[c_matricula],
        "Ubicacion": df[c_ubicacion],
        "Vel_kmh": df[c_vel].apply(_parsear_velocidad),
        "Km_dia": pd.to_numeric(df[c_km], errors="coerce").fillna(0.0),
    })

    if c_ultimo_reporte is not None:
        out["Ultimo_Reporte_Texto"] = df[c_ultimo_reporte].astype(str).str.strip()
        out["Ultimo_Reporte_Seg"] = out["Ultimo_Reporte_Texto"].apply(_parsear_duracion_segundos)
    else:
        log_fn("ℹ El reporte no trae columna 'Último Reporte' -- se omite el chequeo de 'UT pegada'.")
        out["Ultimo_Reporte_Texto"] = None
        out["Ultimo_Reporte_Seg"] = None

    return out


def _parsear_velocidad(texto):
    if texto is None:
        return 0.0
    m = re.search(r"[\d.,]+", str(texto))
    if not m:
        return 0.0
    return float(m.group(0).replace(",", "."))


def _parsear_duracion_segundos(texto):
    """Convierte '3s' / '1m, 49s' / '3m, 5s' / '1h, 5m, 38s' / '1d, 16m, 47s'
    a segundos totales. Mismo criterio documentado en
    ultimas_ut_movimiento.py (_parsear_duracion_segundos): d=86400, h=3600,
    m=60, s=1. Retorna None si el texto no calza con el patrón esperado."""
    if not texto or str(texto).strip().lower() in ("nan", "none", ""):
        return None
    texto = str(texto)
    unidades = {"d": 86400, "h": 3600, "m": 60, "s": 1}
    total = 0
    encontrado = False
    for cantidad, unidad in re.findall(r"(\d+)\s*([dhms])", texto, flags=re.IGNORECASE):
        total += int(cantidad) * unidades[unidad.lower()]
        encontrado = True
    return total if encontrado else None


# ---------------------------------------------------------------------------
# Lectura del Excel de Disponibilidad -- calamine (stylesheet inválido,
# igual que documenta Herramienta 1 / Herramienta 7), detectando encabezado
# en las primeras 3 filas y tolerando layouts distintos por hoja
# (incluye el caso Ocm, donde Ruta y Disponibilidad están en orden invertido
# respecto al resto de las hojas)
# ---------------------------------------------------------------------------
_NOMBRES_ALIAS = ("ut", "ut2", "unidad")
_NOMBRES_ESTATUS = ("estatus", "status", "disponibilidad")
_NOMBRES_CTS = ("cts",)
_NOMBRES_RUTA = ("ruta", "ubicacion / ruta / serv. esp", "servicio")


def _formatear_alias_numerico(valor) -> str:
    """
    Corrige una fragilidad real de python-calamine (16/09/2026, ver el bug
    de la unidad 0140 en la hoja Ocm de Actualizar Desconexión): si la
    columna de alias de una hoja de Disponibilidad es TOTALMENTE numérica
    (ningún alias tipo "R-XXX" que la obligue a quedar como texto),
    calamine la lee como número y pierde los ceros a la izquierda ("0140"
    llega como 140.0). Se reconstruye como texto de 4 dígitos —formato fijo
    de alias numérico usado en todo el proyecto— cuando el valor llega como
    int/float; si ya viene como texto, se retorna tal cual.
    """
    if isinstance(valor, bool):
        return str(valor)
    if isinstance(valor, (numbers.Integral, numbers.Real)):
        try:
            entero = int(valor)
            if entero == valor:
                return f"{entero:04d}"
        except (ValueError, OverflowError):
            pass
    return str(valor)


def leer_disponibilidad(ruta_excel, log_fn=print):
    xl = pd.ExcelFile(ruta_excel, engine="calamine")
    resultado = {}  # {ALIAS_UPPER: {"status":..., "cts":..., "ruta":..., "hoja":...}}

    for hoja in xl.sheet_names:
        hoja_norm = hoja.strip().lower()
        if hoja_norm in HOJAS_SIN_DATOS_UTILES:
            log_fn(f"ℹ Hoja '{hoja}' omitida (sin formato de columnas utilizable).")
            continue

        crudo = xl.parse(hoja, header=None, nrows=3)
        if crudo.empty:
            log_fn(f"ℹ Hoja '{hoja}' vacía, omitida.")
            continue

        fila_header = None
        col_alias = col_estatus = col_cts = col_ruta = None
        for i in range(min(3, len(crudo))):
            fila = [str(v).strip().lower() if pd.notna(v) else "" for v in crudo.iloc[i]]
            ca = next((j for j, v in enumerate(fila) if v in _NOMBRES_ALIAS or "numeracion" in v), None)
            ce = next((j for j, v in enumerate(fila) if any(v.startswith(n) for n in _NOMBRES_ESTATUS)), None)
            if ca is not None and ce is not None:
                fila_header = i
                col_alias, col_estatus = ca, ce
                col_cts = next((j for j, v in enumerate(fila) if v in _NOMBRES_CTS), None)
                col_ruta = next((j for j, v in enumerate(fila) if any(v.startswith(n) for n in _NOMBRES_RUTA)), None)
                break

        if fila_header is None:
            log_fn(f"ℹ Hoja '{hoja}' no calza con ningún formato reconocido, omitida.")
            continue

        df = xl.parse(hoja, header=fila_header)
        df.columns = [str(c).strip().lower() for c in df.columns]
        cols = list(df.columns)

        c_alias = cols[col_alias]
        c_estatus = cols[col_estatus]
        c_cts = cols[col_cts] if col_cts is not None else None
        c_ruta = cols[col_ruta] if col_ruta is not None else None

        for _, fila in df.iterrows():
            alias_raw = fila.get(c_alias)
            if pd.isna(alias_raw) or str(alias_raw).strip() == "":
                continue
            alias = _formatear_alias_numerico(alias_raw).strip().upper()
            # Nota (corregido 16/09/2026): calamine SOLO preserva los ceros a
            # la izquierda cuando la columna es de tipo texto (mezclada con
            # algún alias no numérico como "R-XXX"). Si una hoja llega a ser
            # enteramente numérica (ver el caso real de la unidad 0140 en la
            # hoja Ocm de Actualizar Desconexión), calamine la lee como
            # número y los pierde — _formatear_alias_numerico() los
            # reconstruye. Acá solo limpiamos espacios/puntos sueltos.
            alias = alias.lstrip(".").strip()

            status_raw = fila.get(c_estatus)
            status = str(status_raw).strip() if pd.notna(status_raw) else ""
            status_norm = status.strip().capitalize() if status else ""

            cts_raw = fila.get(c_cts) if c_cts else None
            cts = str(cts_raw).strip() if pd.notna(cts_raw) else ""

            ruta_raw = fila.get(c_ruta) if c_ruta else None
            ruta = str(ruta_raw).strip() if pd.notna(ruta_raw) else ""

            resultado[alias] = {
                "status": status_norm,
                "cts": cts,
                "ruta": ruta,
                "hoja": hoja,
            }

    return resultado


def normalizar_alias(alias):
    """El Alias del Reporte de Estatus Actual a veces trae el prefijo
    'UT: 0491' (visto con datos reales) en vez del alias plano '0491' que
    usa el Excel de Disponibilidad. Se quita ese prefijo para poder cruzar
    ambas fuentes; los alias tipo 'R-020' se dejan intactos.

    También se quitan caracteres sueltos de markdown ('*', '_', '`') que
    a veces vienen pegados al alias en los datos reales (ej. 'R-001*') --
    sin esto, el asterisco suelto rompe el negrita del mensaje de Telegram
    (deja el '*' sin cerrar para el resto del mensaje)."""
    if alias is None:
        return ""
    texto = str(alias).strip()
    texto = re.sub(r"(?i)^\s*ut\s*[:.\-]?\s*", "", texto).strip()
    texto = re.sub(r"[*_`]", "", texto).strip()
    return texto


def _normalizar_para_exclusion(alias):
    """Forma canónica de un alias para comparar contra la lista de unidades a
    obviar (nuevo v3.9): 'r66'/'R-66'/'R-066' -> 'R-066'; '343'/'0343' -> '0343'."""
    a = normalizar_alias(alias).upper().replace(" ", "")
    m = re.fullmatch(r"R-?(\d{1,3})", a)
    if m:
        return f"R-{int(m.group(1)):03d}"
    m = re.fullmatch(r"\d{1,4}", a)
    if m:
        return a.zfill(4)
    return a


def get_disponibilidad(disp_db, alias):
    if not disp_db:
        return None
    return disp_db.get(normalizar_alias(alias).upper())


PATIOS_KEY = "PATIOS"  # clave reservada en rutas_db.json para Patio de los Ejes
ALIAS_RUTAS_KEY = "ALIAS_RUTAS"  # clave reservada: {EJE: {nombre_reportado: nombre_correcto}}
# Claves reservadas nuevas (22/09/2026): ubicaciones donde estar detenido o
# fuera de ruta NO es una anomalía, igual que un Patio.
ABASTECIMIENTO_KEY = "ABASTECIMIENTO"        # bombas / puntos de abastecimiento
SEDES_KEY = "SEDES"                          # sedes (Proveeduría, Transmiranda...)
PUNTOS_RESGUARDO_KEY = "PUNTOS_RESGUARDO"    # {ALIAS_UNIDAD: [direcciones]} -- solo para ESA unidad


# ---------------------------------------------------------------------------
# Registro interno de Rutas y su recorrido correcto
# ---------------------------------------------------------------------------
def cargar_rutas_db(ruta_json):
    with open(ruta_json, encoding="utf-8") as f:
        crudo = json.load(f)

    patios_crudos = crudo.pop(PATIOS_KEY, [])
    alias_crudos = crudo.pop(ALIAS_RUTAS_KEY, {})
    abastec_crudos = crudo.pop(ABASTECIMIENTO_KEY, [])
    sedes_crudas = crudo.pop(SEDES_KEY, [])
    resguardo_crudo = crudo.pop(PUNTOS_RESGUARDO_KEY, {})

    # normalizar cada dirección para comparación rápida
    normalizado = {}
    for eje, rutas in crudo.items():
        normalizado[eje] = {}
        for nombre_ruta, direcciones in rutas.items():
            normalizado[eje][nombre_ruta] = [
                (d, _normalizar_comparacion(limpiar_ubicacion(d))) for d in direcciones
            ]

    normalizado[PATIOS_KEY] = [
        (d, _normalizar_comparacion(limpiar_ubicacion(d))) for d in patios_crudos
    ]

    normalizado[ABASTECIMIENTO_KEY] = [
        (d, _normalizar_comparacion(limpiar_ubicacion(d))) for d in abastec_crudos
    ]
    normalizado[SEDES_KEY] = [
        (d, _normalizar_comparacion(limpiar_ubicacion(d))) for d in sedes_crudas
    ]
    # {ALIAS_UPPER: [direcciones normalizadas]} -- resguardo propio de cada unidad
    normalizado[PUNTOS_RESGUARDO_KEY] = {
        alias.strip().upper(): [_normalizar_comparacion(limpiar_ubicacion(d)) for d in dirs]
        for alias, dirs in resguardo_crudo.items()
    }

    # {EJE: {nombre_reportado_normalizado: nombre_correcto}} -- para resolver
    # "el nombre correcto que debe aparecer en el reporte" cuando Disponibilidad
    # trae un nombre de ruta distinto (o una de varias variantes equivalentes)
    normalizado[ALIAS_RUTAS_KEY] = {
        eje: {_normalizar_comparacion(alias): correcto for alias, correcto in mapeo.items()}
        for eje, mapeo in alias_crudos.items()
    }
    return normalizado


def es_ubicacion_patio(ubicacion_actual, rutas_db, alias=None):
    """True si la ubicación coincide (tolerante a mayúsculas/tildes/espacios
    extra/punto final) con un Patio de los Ejes, un Punto de Abastecimiento
    o una Sede (aplican a cualquier unidad) o -- si se pasa alias -- con el
    Punto de Resguardo documentado PARA ESA unidad. Parámetro alias nuevo
    (22/09/2026), al final y con valor por defecto para no romper llamadas
    existentes."""
    ubic_norm = _normalizar_comparacion(limpiar_ubicacion(ubicacion_actual))
    if not ubic_norm:
        return False
    for clave in (PATIOS_KEY, ABASTECIMIENTO_KEY, SEDES_KEY):
        if any(ubic_norm == norm for _, norm in (rutas_db.get(clave) or [])):
            return True
    if alias:
        propios = (rutas_db.get(PUNTOS_RESGUARDO_KEY) or {}).get(str(alias).strip().upper(), [])
        if ubic_norm in propios:
            return True
    return False


UMBRAL_SIMILITUD_RUTA = 0.90  # alto a propósito: confirmado que rutas distintas
# pueden compartir la mayoría del nombre (ej. "Altagracia de la Montaña - Cua"
# vs "Altagracia de la Montaña - Los Teques" son DOS rutas reales y diferentes)


def _sin_espacios(texto):
    return re.sub(r"\s+", "", texto)


def _slug(texto):
    """Solo letras/números, sin guiones/puntuación/espacios -- para comparar
    nombres de ruta ignorando cómo se haya escrito la puntuación."""
    return re.sub(r"[^a-z0-9]", "", texto)


def _palabras(texto):
    """Solo letras/números separados por un espacio (ignora puntuación)."""
    return re.sub(r"[^a-z0-9]+", " ", texto).strip()


def _buscar_por_prefijo(rutas_del_eje, alias_del_eje, objetivo):
    """Si el texto empieza (palabra por palabra, ignorando puntuación) con el
    nombre de una ruta o alias documentado, devuelve esa ruta. Gana el nombre
    más largo. Exige que el prefijo termine en límite de palabra y tenga al
    menos 8 letras/números, para no emparejar por casualidad."""
    obj = _palabras(objetivo)
    mejor = None  # (largo, nombre_correcto)
    conocidos = [(_palabras(_normalizar_comparacion(n)), n) for n in rutas_del_eje]
    if alias_del_eje:
        conocidos += [(_palabras(a), c) for a, c in alias_del_eje.items() if c in rutas_del_eje]
    for k, correcto in conocidos:
        if len(k.replace(" ", "")) < 8:
            continue
        if obj == k or obj.startswith(k + " "):
            if mejor is None or len(k) > mejor[0]:
                mejor = (len(k), correcto)
    if mejor:
        return mejor[1], rutas_del_eje[mejor[1]]
    return None


def _buscar_en_alias(alias_del_eje, texto_norm, texto_slug):
    """Busca en el mapeo ALIAS_RUTAS por texto exacto o por 'slug'."""
    directo = alias_del_eje.get(texto_norm)
    if directo:
        return directo
    if texto_slug:
        for k, v in alias_del_eje.items():
            if _slug(k) == texto_slug:
                return v
    return None


def _candidatos_truncados(objetivo):
    """Candidatos de nombre de ruta tomando solo lo que va al INICIO del texto
    (ya normalizado a minúsculas/sin tildes), antes de 'apoyo' (o apoya/apoyan),
    de '/' o de '-'. Del más largo al más corto. Se descartan candidatos muy
    cortos (< 4 letras/números) para no emparejar basura."""
    cortes = []
    m = re.search(r"\bapoy\w*", objetivo)
    if m and m.start() > 0:
        cortes.append(objetivo[:m.start()])
    if "/" in objetivo:
        cortes.append(objetivo.split("/", 1)[0])
    base = list(cortes) + [objetivo]
    for b in base:
        for mm in reversed(list(re.finditer(r"-", b))):
            cortes.append(b[:mm.start()])
    vistos, salida = set(), []
    for c in cortes:
        c = re.sub(r"\s+y$", "", c.strip(" -/,.:;()"))
        c = c.strip(" -/,.:;()")
        if len(_slug(c)) >= 4 and c not in vistos and c != objetivo:
            vistos.add(c)
            salida.append(c)
    return salida


def _buscar_ruta_por_nombre(rutas_del_eje, nombre_ruta, alias_del_eje=None):
    """Empareja el nombre de Ruta que trae Disponibilidad contra el nombre
    de ruta documentado.

    Orden de búsqueda:
      0) mapeo EXPLÍCITO de alias (rutas_db.json -> ALIAS_RUTAS), cargado a
         partir del registro de rutas donde el usuario documentó
         "Nombre reportado / Nombre correcto" o variantes equivalentes con
         "=". Esto es lo más confiable porque viene confirmado a mano, no
         inferido por parecido de texto.
      1) coincidencia exacta ignorando mayúsculas/tildes/espacios/guiones.
      2) como último recurso, similitud de texto con un umbral muy alto.

    IMPORTANTE (confirmado con datos reales): rutas TOTALMENTE distintas
    pueden compartir casi todo el nombre salvo el sufijo final (ej.
    'Altagracia de la Montaña - Cua' vs 'Altagracia de la Montaña - Los
    Teques'). Por eso NUNCA se usa 'contains' parcial de un nombre dentro de
    otro para nombres de ruta -- solo coincidencia exacta, alias explícito,
    o similitud con umbral muy alto (ver UMBRAL_SIMILITUD_RUTA).
    """
    if not nombre_ruta:
        return None
    objetivo = _normalizar_comparacion(nombre_ruta)
    if not objetivo:
        return None

    objetivo_slug = _slug(objetivo)

    # 0) alias explícito confirmado por el usuario (exacto, o igual ignorando
    #    puntuación/espacios -- así "MARICHE-PETARE" y "Mariche - Petare" dan igual)
    if alias_del_eje:
        nombre_correcto = _buscar_en_alias(alias_del_eje, objetivo, objetivo_slug)
        if nombre_correcto and nombre_correcto in rutas_del_eje:
            return nombre_correcto, rutas_del_eje[nombre_correcto]

    # 1) coincidencia exacta, ignorando guiones/puntuación/espacios
    for nombre_doc, direcciones in rutas_del_eje.items():
        if _slug(_normalizar_comparacion(nombre_doc)) == objetivo_slug:
            return nombre_doc, direcciones

    # 1a) NUEVO (22/09/2026): el texto EMPIEZA con un nombre de ruta/alias
    #     conocido y luego sigue una nota (ej. nombre largo de Charallave Norte
    #     pegado con otros tramos). Se elige el nombre conocido más largo.
    pref = _buscar_por_prefijo(rutas_del_eje, alias_del_eje, objetivo)
    if pref is not None:
        return pref

    # 1b) NUEVO (22/09/2026): si el texto completo no coincidió con nada, se
    #     prueba solo con el nombre que viene AL INICIO, antes de "apoyo",
    #     "/" o "-" (ej. "Higuerote - Caucagua apoya Petare" -> "Higuerote -
    #     Caucagua"). Solo es respaldo: el texto completo siempre se prueba
    #     primero, así las rutas reales con guiones internos no se recortan.
    for candidato in _candidatos_truncados(objetivo):
        cand_slug = _slug(candidato)
        if alias_del_eje:
            nombre_correcto = _buscar_en_alias(alias_del_eje, candidato, cand_slug)
            if nombre_correcto and nombre_correcto in rutas_del_eje:
                return nombre_correcto, rutas_del_eje[nombre_correcto]
        for nombre_doc, direcciones in rutas_del_eje.items():
            if _slug(_normalizar_comparacion(nombre_doc)) == cand_slug:
                return nombre_doc, direcciones

    # 2) similitud de texto -- SOLO como último recurso, umbral alto para
    #    typos menores (NO para nombres con un sufijo distinto)
    mejor_nombre, mejor_ratio = None, 0.0
    for nombre_doc in rutas_del_eje:
        ratio = difflib.SequenceMatcher(None, objetivo, _normalizar_comparacion(nombre_doc)).ratio()
        if ratio > mejor_ratio:
            mejor_nombre, mejor_ratio = nombre_doc, ratio
    # también contra los alias documentados (ej. "27 FEBRERO- LOS DOS CAMINOS"
    # vs el alias "27 DE FEBRERO - LOS DOS CAMINOS") -- mismo umbral alto
    if alias_del_eje:
        for alias_norm, correcto in alias_del_eje.items():
            if correcto not in rutas_del_eje:
                continue
            ratio = difflib.SequenceMatcher(None, objetivo, alias_norm).ratio()
            if ratio > mejor_ratio:
                mejor_nombre, mejor_ratio = correcto, ratio
    if mejor_nombre is not None and mejor_ratio >= UMBRAL_SIMILITUD_RUTA:
        return mejor_nombre, rutas_del_eje[mejor_nombre]

    return None


def _clave_rutas_db_para_eje(eje):
    """El registro de rutas agrupa 'Serv Cont' y 'Serv Esp' bajo un único
    eje 'SERV' (Servicios Especiales) -- ambos mapean a la misma clave."""
    if eje in ("Serv Cont", "Serv Esp"):
        return "SERV"
    return eje.upper()


def resolver_nombre_ruta(eje, nombre_original, rutas_db):
    """Devuelve el nombre CORRECTO de la ruta para mostrar en los reportes
    (nuevo v3.7). Si nombre_original coincide con una ruta conocida (por
    alias explícito, coincidencia exacta, o similitud alta), devuelve el
    nombre documentado como correcto. Si no coincide con nada conocido
    (ej. una nota de texto libre como "FALLA POR CROCHET"), devuelve el
    texto original tal cual, sin inventar nada."""
    if not nombre_original:
        return nombre_original
    clave = _clave_rutas_db_para_eje(eje)
    rutas_del_eje = rutas_db.get(clave)
    if not rutas_del_eje:
        return nombre_original
    alias_del_eje = rutas_db.get(ALIAS_RUTAS_KEY, {}).get(clave)
    match = _buscar_ruta_por_nombre(rutas_del_eje, nombre_original, alias_del_eje)
    return match[0] if match else nombre_original


def es_fuera_de_ruta(eje, nombre_ruta_asignada, ubicacion_actual, rutas_db):
    """
    Devuelve:
      True  -> la unidad está fuera de la ruta que tiene asignada
      False -> la ubicación coincide con la ruta asignada
      None  -> no se puede evaluar (eje sin registro de rutas todavía,
               o unidad sin ruta asignada) -- se omite el chequeo sin error
    """
    rutas_del_eje = rutas_db.get(_clave_rutas_db_para_eje(eje))
    if not rutas_del_eje:
        return None  # eje sin rutas documentadas aún
    if not nombre_ruta_asignada or nombre_ruta_asignada.strip().upper() in _RUTA_NO_ASIGNADA:
        return None  # unidad sin ruta/servicio asignado, no aplica el chequeo

    match = _buscar_ruta_por_nombre(rutas_del_eje, nombre_ruta_asignada,
                                     rutas_db.get(ALIAS_RUTAS_KEY, {}).get(_clave_rutas_db_para_eje(eje)))
    if match is None:
        return None  # la ruta asignada no está en el registro todavía

    _, direcciones_normalizadas = match
    if not direcciones_normalizadas:
        return None  # ruta ya registrada pero SIN direcciones cargadas todavía
        # (ej. ruta nueva confirmada, pendiente de que se entregue su recorrido)
        # -- se omite el chequeo en vez de marcar "Fuera de ruta" por defecto.

    ubic_norm = _normalizar_comparacion(limpiar_ubicacion(ubicacion_actual))
    if not ubic_norm:
        return None

    for _, direccion_norm in direcciones_normalizadas:
        if ubic_norm in direccion_norm or direccion_norm in ubic_norm:
            return False  # coincide, está en su ruta
    return True  # no coincidió con ninguna dirección válida de su ruta


# ---------------------------------------------------------------------------
# Formato de mensaje (idéntico al de Excesos de Velocidad / Sin Información,
# con las 3 variantes de Incidencia pedidas)
# ---------------------------------------------------------------------------
def _encabezado_eje(eje, subeje):
    if eje == "Serv Cont":
        return "Eje *Serv* *Cont.*"
    if eje == "Serv Esp":
        return "Eje *Serv* *Esp.*"
    if subeje:
        return f"Eje *{eje}* - Sub eje *{subeje}*"
    return f"Eje *{eje}*"


def _formato_mensaje(saludo, eje, subeje, unidad, vin, cts, cuerpo_incidencia):
    lineas = [
        saludo,
        f"* {_encabezado_eje(eje, subeje)}",
        f"* Unidad *{unidad}* - Vin *{vin}*",
        f"* CTS: *{cts}*",
        "",
        cuerpo_incidencia,
    ]
    return "\n".join(lineas)


def get_vin(vin_db, alias):
    """vin_db=None -> usa el VIN_DB compartido real del proyecto
    (tools.excesos_pdf.VIN_DB); se puede pasar uno explícito para pruebas."""
    if vin_db is None:
        return get_vin_compartido(alias)
    return vin_db.get(str(alias).strip().upper(), "S.I.")


# ---------------------------------------------------------------------------
# Orquestador principal
# ---------------------------------------------------------------------------
ORDEN_EJES = ["Am", "Blv", "Met", "Ocm", "Pz", "Serv Cont", "Serv Esp"]
ORDEN_INCIDENCIAS = ["Fuera de ruta", "Ruta no reconocida", "Movimiento no reportado",
                     "Sin información", "Exceso de velocidad", "Sin movimiento", "UT pegada"]


def generar_reportes_ojo_de_halcon(ruta_estatus, ruta_disponibilidad, ruta_rutas_json,
                                    vin_db=None, log_fn=print, excluir_unidades=None):
    """
    Retorna list[str]: un mensaje por cada incidencia detectada, organizados
    por Eje y Sub-eje (orden fijo: Am, Blv, Met, Ocm, Pz, Serv Cont, Serv
    Esp -- mismo orden que usa ultimas_ut_movimiento.py) y, dentro de cada
    Eje/Sub-eje, en este orden de incidencia: Fuera de ruta, Movimiento no
    reportado, Exceso de velocidad, Sin movimiento, UT pegada (nuevo v3.3).

    Solo se evalúan unidades que en Disponibilidad figuren "Operativa", o
    "Inoperativa" con Km del día > 1.0 -- el resto de las Inoperativas se
    omite por completo (no generan ninguna incidencia, nuevo v3.3). Las
    unidades sin registro en Disponibilidad también se omiten, porque no se
    puede confirmar que cumplan ese criterio.

    vin_db=None (por defecto) usa el VIN_DB compartido real del proyecto
    (tools.excesos_pdf.VIN_DB / tools.sin_informacion.py).

    excluir_unidades (nuevo v3.9): set opcional de alias (en cualquier
    formato reconocible, ej. {"R-036", "0140"}) que se omiten de TODAS las
    incidencias por pedido explícito del usuario -- paso opcional que el
    bot pregunta antes de generar el reporte.
    """
    excluir_norm = {_normalizar_para_exclusion(u) for u in excluir_unidades} if excluir_unidades else None
    df = leer_estatus_actual(ruta_estatus, log_fn=log_fn)
    disp_db = leer_disponibilidad(ruta_disponibilidad, log_fn=log_fn)
    rutas_db = cargar_rutas_db(ruta_rutas_json)

    saludo = saludo_actual()
    entradas = []  # (indice_eje, subeje_o_vacio, indice_incidencia, mensaje)
    omitidas_sin_disponibilidad = 0
    omitidas_por_status = 0
    omitidas_manual = 0

    for _, fila in df.iterrows():
        alias = normalizar_alias(fila["Alias"])
        if not alias:
            continue
        if excluir_norm and _normalizar_para_exclusion(alias) in excluir_norm:
            omitidas_manual += 1
            continue  # unidad obviada por pedido explícito del usuario
        info_eje = extraer_info_eje(fila["Matricula"], log_fn=log_fn)
        if info_eje is None:
            continue
        eje, subeje = info_eje["eje"], info_eje["subeje"]

        disp = get_disponibilidad(disp_db, alias)
        km_dia = fila["Km_dia"]
        idx_eje = ORDEN_EJES.index(eje) if eje in ORDEN_EJES else len(ORDEN_EJES)
        clave_orden = (idx_eje, subeje or "")

        if disp is None:
            omitidas_sin_disponibilidad += 1
            # Excepción nueva v3.6: si la unidad SÍ tuvo movimiento real
            # (Km del día > 1.0) pero no aparece en Disponibilidad, se
            # reporta "Sin información" en vez de omitirse en silencio --
            # la idea es detectar unidades con movimiento que quedaron
            # fuera del registro de Disponibilidad.
            if km_dia > UMBRAL_KM_DIA_SIN_REPORTE:
                vin = get_vin(vin_db, alias)
                ubicacion = fila["Ubicacion"]
                cuerpo = (
                    "*Incidencia:* Sin información\n"
                    f"*Observación:* No se ha recibido información *({km_dia:.2f} km recorridos).* "
                    f"y se visualiza en *{limpiar_ubicacion(ubicacion)}.*"
                )
                idx_inc = ORDEN_INCIDENCIAS.index("Sin información")
                msg = _formato_mensaje(saludo, eje, subeje, alias, vin, "S.I", cuerpo)
                entradas.append((clave_orden, idx_inc, msg))
            continue  # sin registro en Disponibilidad -> no se evalúa el resto de incidencias

        status = disp["status"]

        # Filtro nuevo v3.3: solo Operativa, o Inoperativa con Km del día > 1.0
        if status == "Operativa":
            pass
        elif status == "Inoperativa" and km_dia > UMBRAL_KM_DIA_SIN_REPORTE:
            pass
        else:
            omitidas_por_status += 1
            continue

        cts = formatear_cts_ojo(disp["cts"])
        ruta_asignada = disp["ruta"]

        vin = get_vin(vin_db, alias)
        ubicacion = fila["Ubicacion"]
        vel = fila["Vel_kmh"]
        ultimo_reporte_seg = fila["Ultimo_Reporte_Seg"]
        ultimo_reporte_texto = fila["Ultimo_Reporte_Texto"]

        def agregar(tipo_incidencia, cuerpo):
            idx_inc = ORDEN_INCIDENCIAS.index(tipo_incidencia)
            msg = _formato_mensaje(saludo, eje, subeje, alias, vin, cts, cuerpo)
            entradas.append((clave_orden, idx_inc, msg))

        # 1) Exceso de velocidad
        if vel >= UMBRAL_VELOCIDAD_KMH:
            cuerpo = (
                "*Incidencia:* Exceso de velocidad \n"
                f"*Incumplimiento:* *{int(round(vel))} km/h* en *autopista*"
            )
            agregar("Exceso de velocidad", cuerpo)

        # 2) Movimiento no reportado (Inoperativa + Km del día > 1.0)
        if status == "Inoperativa" and km_dia > UMBRAL_KM_DIA_SIN_REPORTE:
            cuerpo = (
                "*Incidencia:* Movimiento no reportado \n"
                f"*Observación:* *Falla* en movimiento *({km_dia:.2f} km recorridos)*"
            )
            agregar("Movimiento no reportado", cuerpo)

        # 3) Fuera de ruta -- primero se chequea Patio de los Ejes (nuevo v3.2):
        # estar en Patio mientras se tiene una ruta REAL Y RECONOCIDA asignada
        # ES "Fuera de ruta", con la Observación fija "Se visualiza en Patio"
        # en vez de la dirección real.
        #
        # IMPORTANTE (bug real corregido con datos reales): la columna
        # Ruta/Servicio de Disponibilidad no siempre trae un nombre de ruta --
        # a veces trae notas de texto libre sobre por qué la unidad no está
        # circulando (ej. "FALLA POR CROCHET", "resguardo", "se dirige a
        # buscar repuesto", "Disposición de la Gerencia"). Generar "Fuera de
        # ruta" para esos casos sería un falso positivo: la unidad no está
        # incumpliendo una ruta, está fuera de servicio por otra razón. Por
        # eso el chequeo de Patio solo dispara si esa Ruta coincide con una
        # ruta RECONOCIDA en el registro (rutas_db) -- si no coincide con
        # ninguna ruta conocida, se omite (no se puede verificar con certeza).
        en_patio = es_ubicacion_patio(ubicacion, rutas_db, alias)
        ruta_asignada_valida = bool(ruta_asignada) and ruta_asignada.strip().upper() not in _RUTA_NO_ASIGNADA
        rutas_del_eje = rutas_db.get(_clave_rutas_db_para_eje(eje)) if ruta_asignada_valida else None
        ruta_reconocida = None
        if ruta_asignada_valida and rutas_del_eje:
            alias_del_eje = rutas_db.get(ALIAS_RUTAS_KEY, {}).get(_clave_rutas_db_para_eje(eje))
            match_ruta = _buscar_ruta_por_nombre(rutas_del_eje, ruta_asignada, alias_del_eje)
            if match_ruta is not None:
                ruta_reconocida = match_ruta[0]

        # 3b) Ruta no reconocida (nuevo, pedido explícito del usuario): la
        # unidad SÍ tiene una Ruta asignada en Disponibilidad (no es
        # S.R/S-OP/S/OP/S.S/vacío) y el Eje SÍ tiene rutas cargadas en el
        # registro (rutas_db.json), pero ese texto de Ruta no coincide con
        # ninguna ruta conocida (ni por alias explícito, ni exacta, ni por
        # similitud alta). Antes esto se omitía en silencio (mismo criterio
        # que "eje sin rutas registradas"); ahora se avisa con la UT y el
        # texto de Ruta tal cual viene en Disponibilidad, para poder cargarla
        # al registro (o confirmar que es una nota operativa y no una ruta
        # real, ej. "FALLA POR CROCHET", "SS", "Disposición de la Gerencia").
        # Formato simple pedido explícitamente, SIN el encabezado habitual
        # de Eje/Unidad/Vin/CTS que usan las demás incidencias.
        if ruta_asignada_valida and rutas_del_eje and ruta_reconocida is None and not _es_nota_operativa(ruta_asignada):
            idx_inc = ORDEN_INCIDENCIAS.index("Ruta no reconocida")
            msg_ruta_no_reconocida = f"Ut :{alias}, ruta : {ruta_asignada} (Ruta no reconocida)"
            entradas.append((clave_orden, idx_inc, msg_ruta_no_reconocida))

        # Nombre a MOSTRAR en los reportes: el "nombre correcto" documentado
        # si la Ruta asignada coincide con alguno conocido (nuevo v3.7); si no
        # coincide con nada conocido (ej. una nota de texto libre), se deja
        # el texto original tal cual, sin inventar nada.
        ruta_para_mostrar = resolver_nombre_ruta(eje, ruta_asignada, rutas_db) if ruta_asignada else ruta_asignada

        if en_patio and ruta_reconocida:
            cuerpo = (
                "*Incidencia:* Fuera de ruta\n"
                f" *Ruta:* {ruta_reconocida}\n"
                "*Observación:* Se visualiza en *Patio*"
            )
            agregar("Fuera de ruta", cuerpo)
            fuera_de_ruta_bool = True
        elif not en_patio:
            fuera_de_ruta_bool = es_fuera_de_ruta(eje, ruta_asignada, ubicacion, rutas_db) is True
            if fuera_de_ruta_bool:
                cuerpo = (
                    "*Incidencia:* Fuera de ruta\n"
                    f" *Ruta:* {ruta_para_mostrar}\n"
                    f"*Observación:* Se visualiza en *{limpiar_ubicacion(ubicacion)}.*"
                )
                agregar("Fuera de ruta", cuerpo)
        else:
            fuera_de_ruta_bool = False

        # 4) Validación Parada / "UT sin movimiento" (reglas v3.4, techo
        # agregado en v3.5): solo se genera si la unidad está detenida
        # (Vel = 0) Y:
        #   - el 'Último Reporte' lleva MÁS de 30 minutos parada, sin importar
        #     si está o no en su ruta, O
        #   - lleva MÁS de 10 minutos parada Y además está Fuera de ruta.
        # Y ADEMÁS, en cualquiera de los dos casos, lleva COMO MÁXIMO 12
        # horas paradas -- más de 12h se omite (mismo techo que "UT pegada").
        # Las unidades detenidas en Patio se OMITEN siempre de este reporte
        # (estar parada en Patio no es una anomalía).
        if (vel == 0 and not en_patio and ultimo_reporte_seg is not None
                and ultimo_reporte_seg <= UMBRAL_SIN_MOVIMIENTO_MAX_SEG):
            dispara = (
                ultimo_reporte_seg > UMBRAL_SIN_MOVIMIENTO_SEG
                or (ultimo_reporte_seg > UMBRAL_SIN_MOVIMIENTO_FUERA_RUTA_SEG and fuera_de_ruta_bool)
            )
            if dispara:
                duracion_legible = _formatear_duracion_legible(ultimo_reporte_seg)
                cuerpo = (
                    "Validación Parada\n"
                    f"Ruta: {ruta_para_mostrar or 'S.I.'}\n"
                    "Ubicación de Parada:\n"
                    f"{limpiar_ubicacion(ubicacion)}. ({duracion_legible} aproximadamente)\n"
                    "¿Nos podrían indicar si hay algún inconveniente con la UT?"
                )
                agregar("Sin movimiento", cuerpo)

        # 5) UT pegada (nuevo v3.2, techo agregado en v3.3): columna 'Último
        # Reporte' > 1 hora Y <= 12 horas. Más de 12h se omite (se asume que
        # ya hay otra gestión encargándose de esas unidades). Se omite sola
        # si el archivo no trae esa columna (Ultimo_Reporte_Seg queda en
        # None -- ver leer_estatus_actual).
        if ultimo_reporte_seg is not None and \
                UMBRAL_ULTIMO_REPORTE_SEG < ultimo_reporte_seg <= UMBRAL_ULTIMO_REPORTE_MAX_SEG:
            cuerpo = f"*Incidencia:* La siguiente ut tiene ya {ultimo_reporte_texto} pegada"
            agregar("UT pegada", cuerpo)

    entradas.sort(key=lambda e: (e[0], e[1]))
    mensajes = [msg for _, _, msg in entradas]

    if omitidas_por_status:
        log_fn(f"ℹ {omitidas_por_status} unidad(es) omitida(s) por no cumplir el filtro "
               f"Operativa / Inoperativa con Km del día > {UMBRAL_KM_DIA_SIN_REPORTE}.")
    if omitidas_sin_disponibilidad:
        log_fn(f"ℹ {omitidas_sin_disponibilidad} unidad(es) omitida(s) por no tener registro "
               f"en el Excel de Disponibilidad (no se puede confirmar su Estatus).")
    if omitidas_manual:
        log_fn(f"ℹ {omitidas_manual} unidad(es) omitida(s) por pedido explícito del usuario.")

    return mensajes


# ---------------------------------------------------------------------------
# Modo standalone de prueba (NO es el modo de uso final -- el uso final es
# que bot.py importe generar_reportes_ojo_de_halcon() como hacen las otras
# 8 herramientas). Este bloque es solo para poder probar el módulo suelto
# con doble clic en Windows sin que la consola se cierre de inmediato.
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import sys
    import traceback

    try:
        print("=== Ojo de Halcón -- modo de prueba standalone ===\n")
        print("Este script es un MÓDULO de lógica de negocio: la forma final de uso")
        print("es a través del bot de Telegram (bot.py), igual que las otras 8")
        print("herramientas del proyecto. Este modo aquí es solo para probar el")
        print("cruce de los 2 Excel sin depender del bot.\n")

        if len(sys.argv) >= 4:
            ruta_estatus, ruta_disp, ruta_rutas = sys.argv[1], sys.argv[2], sys.argv[3]
        else:
            ruta_estatus = input("Ruta del Excel 'Reporte de Estatus Actual': ").strip('"').strip()
            ruta_disp = input("Ruta del Excel 'Disponibilidad': ").strip('"').strip()
            ruta_rutas = input("Ruta de rutas_db.json (Enter para usar el de esta carpeta): ").strip('"').strip()
            if not ruta_rutas:
                ruta_rutas = "rutas_db.json"

        mensajes = generar_reportes_ojo_de_halcon(
            ruta_estatus, ruta_disp, ruta_rutas, vin_db=None, log_fn=print
        )

        print(f"\n=== {len(mensajes)} mensaje(s) generado(s) ===\n")
        for m in mensajes:
            print("-----")
            print(m)
            print()

    except Exception:
        print("\n*** OCURRIÓ UN ERROR ***\n")
        traceback.print_exc()

    finally:
        input("\nPresioná Enter para cerrar...")

