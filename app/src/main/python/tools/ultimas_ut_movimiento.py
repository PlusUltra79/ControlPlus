"""
Últimas UT en Movimiento — Reporte por Eje (hora manual)
----------------------------------------------------------
Lee el "Reporte de Estatus Actual" del GTRMax (.xls HTML disfrazado) y,
a partir de la columna "Último Reporte" (tiempo transcurrido desde el
último reporte, ej: "12s", "1m, 49s", "3m, 5s"), determina qué unidades
llevan MENOS de 3 minutos sin reportar → están en movimiento.

El Eje de cada unidad se deduce del sub-eje/etiqueta que viene pegado al
VIN en la columna "Matrícula" (ej: "VIN: 4919 Norte" → sub-eje "Norte"),
usando el mapeo SUBEJE_A_EJE. La columna "Ubicación" se recorta para
mostrar solo la calle/sector, sin Parroquia/Municipio/Estado/Venezuela.

Genera un reporte independiente por cada Eje, con el siguiente formato fijo:

  Con unidades en movimiento:
    EJE

    El Centro de Control de Operaciones informa que para la (hora)
    se visualizan las siguientes UT en movimiento:
    1. 0648 Calle Principal con Calle Cumbo. San Fernando del Guapo.
    2. R-050 Calle Zapico. El Delirio. San José de Río Chico.
    ...

  Sin unidades en movimiento:
    EJE

    El Centro de Control de Operaciones informa que para la (hora)
    no se visualizan UT en movimiento.

La hora del encabezado (nuevo — 14 de septiembre de 2026) se normaliza
SIEMPRE a formato de 12 horas con Am/Pm, ej. "08:00 Pm" — sin importar si
el usuario la escribió en 24h ("20:00") o en 12h ("08:00 pm"). Ver
parsear_hora_manual().

Compatible con archivos .xls exportados por el sistema GTR.
Requiere: pip install beautifulsoup4 lxml

Uso: python ultimas_ut_movimiento.py
"""

from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, time as dtime
import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os

# ── Configuración ─────────────────────────────────────────────────────────────

# Orden de salida de los grupos (Eje) en el reporte final.
ORDEN_EJES = ["Am", "Blv", "Met", "Ocm", "Pz", "Serv Cont", "Serv Esp"]

# Mapeo sub-eje/etiqueta (tal como aparece pegado al VIN en "Matrícula")
# → grupo de Eje al que pertenece, según la tabla dada por el usuario.
SUBEJE_A_EJE = {
    "norte": "Am", "sur": "Am", "centro": "Am",
    "acevedo": "Blv", "brion": "Blv", "brión": "Blv",
    "buroz": "Blv", "paez": "Blv", "páez": "Blv", "pedro gual": "Blv",
    "met": "Met", "suc": "Met", "sucre": "Met",
    "tuy i": "Ocm", "tuy ii": "Ocm",
    "pz": "Pz",
    "cont": "Serv Cont", "ggm": "Serv Cont", "cttd": "Serv Cont", "cttd.": "Serv Cont",
    "esp": "Serv Esp", "bus esc": "Serv Esp",
}

# Excepción manual: esta unidad específica va en Serv Esp aunque su
# etiqueta diga "Pz".
UNIDADES_FORZADAS = {
    "R-036": "Serv Esp",
}

# Tiempo máximo (en segundos) desde el "Último Reporte" para considerar
# que la unidad sigue en movimiento.
TOLERANCIA_SEGUNDOS = 3 * 60

# Ubicaciones de "Patio de Vencedores": unidades que transmiten (y por lo
# tanto pasarían la regla de "en movimiento") pero en realidad están
# resguardadas en patio, no circulando. Se comparan ya CORTADAS, en el
# mismo formato que deja limpiar_ubicacion() (sin Parroquia/Municipio/
# Estado/Venezuela), por lo que se omiten del reporte antes de agruparlas.
UBICACIONES_PATIO = {
    # Patio de los Ejes -- extraído de "Recorrido Correcto de las unidades
    # (actualizado)" (27/09/2026). Aplica a CUALQUIER unidad.
    "Calle Los Teques - San Pedro con Local-7, Avenida Víctor Baptista. Urbanización Santo Omero. Los Teques.",
    "Local-7, Avenida Víctor Baptista con Calle Los Teques - San Pedro. Urbanización Santo Omero. Los Teques.",
    "Local-7, Avenida Víctor Baptista con Vía Cárcel de Mujeres. Ramo Verde. Los Teques.",
    "Local-7, Avenida Bicentenaria con Calle Ali Primera. Urbanización El Berbech. Los Teques.",
    "Troncal-9, Carretera Caucagua - El Clavo con Calle El Castaño. Panaquire. Caucagua.",
    "Calle El Calvario entre Vía El Cien y Avenida Manzanares. Maturín. Tacarigua de Mamporal.",
    "Carretera Higuerote - Curiepe con Troncal-12, Carretera Higuerote - Caucagua. Ciudad Balneario Higuerote. Higuerote.",
    "Carretera Tacarigua - Río Chico entre Calle Rincón Bonito y Local-8, Carretera Mamporal - Barlovento. Santo Domingo. Mamporal.",
    "Calle Local 8 con Local-8, Avenida Rafael Arevalo Gonzalez. El Delirio. San José de Río Chico.",
    "Troncal-9, Autopista Cacique Guaicaipuro entre Avenida Principal de La Urbina y Distribuidor La Urbina. La Urbina. Caracas.",
    "Calle 3 entre Calle 2 - 3 y Avenida Principal de La Urbina. La Urbina. Caracas.",
    "Avenida Gumeral entre Calle 7 de Abril y Calle Manga de Coleo. Barrio Campo Elías. Charallave.",
    "Transversal 1 entre Calle 2 y Transversal 3. Sector San José. Guatire.",
    "Calle Principal La Rosa con Calle C. Sector Las Rosas. Guatire.",
}

# Direcciones de la lista ANTERIOR (Patio de Vencedores) que NO aparecen ni en
# "Patio de los Ejes" ni en "Punto de Resguardo" del Word actualizado. Se
# conservan (siguen omitiendo a cualquier unidad) para no cambiar el
# comportamiento sin confirmación; revisar y borrar las que ya no apliquen.
UBICACIONES_PATIO_ANTERIORES = {
    "Avenida Principal El Paso. Urbanización El Paso. Los Teques.",
    "Calle 2 entre Transversal 1 y Avenida Principal Antonio Machado. Sector San José. Guatire.",
    "Calle Chaid Torbay con Calle Las Minas. Rosaleda Norte. San Antonio de Los Altos.",
    "Calle Cumaná. La Raiza. Santa Lucía.",
    "Calle El Bolsillo con Calle Mango de Ocaita. El Guapo. El Guapo.",
    "Calle El Bolsillo con Ramal-35, Calle Real. El Guapo.",
    "Calle José Gregorio Hernández. Panaquire. Caucagua.",
    "Calle Las Mercedes con Calle El Tanque. Corralito. Carrizal.",
    "Calle Los Teques - San Pedro con Local-7, Avenida Víctor Baptista. Urbanización Santo Omero.",
    "Calle Principal La Coromoto con Primera Transversal. San Fernando del Guapo.",
    "Campo Santo. Las Mercedes de Páparo.",
    "Carretera Mamporal - Barlovento entre Avenida Perimetral y Local-8, Calle Miranda. El Delirio. San José de Río Chico.",
    "Carretera Tacarigua - Río Chico entre Calle Araguaney y Calle Rincón Bonito. Santo Domingo. Mamporal.",
    "Eje Rural de Pueblo Nuevo. Campo Santo. Las Mercedes de Páparo.",
    "Empalme Principal El Paso con Avenida Principal El Paso. Barrio José Gregorio Hernández. Los Teques.",
    "La Raiza. Santa Lucía.",
    "Local-6, Carretera Cúa - Tácata. Piñango. Tácata.",
    "Local-7, Avenida Víctor Baptista con Empalme Víctor Baptista. Ramo Verde. Los Teques.",
    "Local-7, Avenida Víctor Baptista entre Empalme Víctor Baptista y Vía Cárcel de Mujeres. Ramo Verde. Los Teques.",
    "Local-8, Avenida Rafael Arevalo Gonzalez entre Calle Local 8 y Calle Bolívar. El Delirio. San José de Río Chico.",
    "Los Teques.",
    "Vía Soapire con Calle El Olivo. Los Guires. Santa Lucía.",
}

# Punto de Resguardo -- POR UNIDAD (27/09/2026): {ALIAS: {direcciones}}. La
# unidad solo se omite si está en SU propio punto de resguardo; la misma
# dirección en otra unidad NO la omite.
PUNTOS_RESGUARDO = {
    "R-047": {
        "Altagracia de La Montaña.",
    },
    "R-010": {
        "Carretera Araira - Yaguapa.",
    },
    "R-049": {
        "El Guapo.",
    },
    "R-050": {
        "Barlovento. Río Chico.",
    },
    "0646": {
        "Vía a Cumbo con Troncal-9, Carretera El Clavo - El Guapo.",
        "Calle Local 8 con Local-8, Avenida Rafael Arevalo Gonzalez. El Delirio. San José de Río Chico.",
    },
    "0648": {
        "Local-8, Carretera Mamporal - Barlovento entre Local-8, Calle Miranda y Recta De Paparo. La Madre Nueva. San José de Río Chico.",
    },
    "0667": {
        "Cúpira. Machurucuto.",
    },
    "R-031": {
        "Hoyo de La Puerta. Caracas.",
    },
    "R-064": {
        "Vía a Turgua. Puerta Negra. Caracas.",
    },
    "0429": {
        "Calle Este 1 entre Avenida Principal Norte 3 y Calle Sector H. El Carpintero - El Limoncito. Caracas.",
    },
    "0495": {
        "Avenida Venezuela con Puente Avenida Las Acacias. Los Caobos. Caracas.",
    },
    "R-025": {
        "Local-4, Carretera Vieja Petare - Santa Lucía con Calle El Colegio.",
    },
    "R-033": {
        "Avenida Principal Turumo entre Calle Razetti y Calle Brisas Bolívar. Turumo. Caracas.",
    },
    "R-051": {
        "Avenida Principal de Manzanares con Redoma Principal de Manzanares. Manzanares. Caracas.",
    },
    "R-039": {
        "Calle Principal Urbanización Loma Real. Zona Industrial. Charallave.",
    },
    "0485": {
        "Calle 4. Barrio 23 de Enero. Ocumare del Tuy.",
    },
    "0678": {
        "Calle 2 Salamanca entre Calle Salamanca y Calle Salamanca 1. Cúa. Cúa.",
        "Avenida Principal. Evelinda. Cúa.",
    },
    "R-042": {
        "Tomuso. El Cartanal.",
    },
    "R-044": {
        "Sta. Rita. Santa Lucía.",
    },
    "R-054": {
        "La Vega. Santa Lucía.",
    },
    "R-055": {
        "El Hornito. Santa Teresa del Tuy.",
        "Autopista A Charallave con Retorno A Charallave. Cabrera. Ocumare del Tuy.",
    },
    "R-063": {
        "Vía a Santa Rita. Sta. Rita. Santa Lucía.",
    },
    "0454": {
        "Local-11, Carretera La Raiza con Empalme La Raiza. Las Dos Lagunas. El Cartanal.",
    },
    "R-016": {
        "Calle Principal del Quemadito con Local-12, Carretera Nacional. Sector Las Rosas. Guatire.",
    },
    "0446": {
        "Carretera Vieja Los Teques - Las Adjuntas. Sector Los Mangos. Los Teques.",
        "Carretera Vieja Los Teques - Las Adjuntas con Local-7, Avenida Bicentenaria. Sector Los Mangos. Los Teques.",
    },
    "0471": {
        "Guarenal. Los Teques.",
    },
    "0592": {
        "Calle Principlal Jabillar. Barrio El Turpial. Los Teques.",
    },
    "R-011": {
        "Ramo Verde. Los Teques.",
    },
    "R-036": {
        "Local-9, Carretera San José de Los Altosl - San Diego con Local-9, Carretera San Diego - San José de Los Altos. Urbanización Colinas de Pasatiempo. San Diego de Los Altos.",
    },
    "0505": {
        "Local-7, Avenida Bicentenaria. Sector Los Mangos. Los Teques.",
        "Local-7, Avenida Bicentenaria con Calle Ali Primera. Urbanización El Berbech. Los Teques.",
    },
    "0562": {
        "Zona El Tamar. Los Teques.",
    },
    "0591": {
        "Barrio Venezuela. Los Teques.",
    },
    "0671": {
        "Calle Principal de Los Manguitos con Calle Cristóbal Conde. Caricuao. Caracas.",
    },
}


def _normalizar_para_comparar(texto: str) -> str:
    """Normaliza espacios/mayúsculas/punto final para comparar ubicaciones
    de forma tolerante a diferencias menores de formato."""
    t = re.sub(r'\s+', ' ', str(texto).strip().lower())
    return t.rstrip('.')


_UBICACIONES_PATIO_NORM = ({_normalizar_para_comparar(u) for u in UBICACIONES_PATIO}
                           | {_normalizar_para_comparar(u) for u in UBICACIONES_PATIO_ANTERIORES})
_PUNTOS_RESGUARDO_NORM = {
    alias.strip().upper(): {_normalizar_para_comparar(d) for d in dirs}
    for alias, dirs in PUNTOS_RESGUARDO.items()
}


def es_ubicacion_patio(ubicacion_limpia: str, alias: str = None) -> bool:
    """True si la ubicación (ya procesada por limpiar_ubicacion) es un Patio
    de los Ejes (cualquier unidad) o -- si se pasa alias -- el Punto de
    Resguardo documentado PARA ESA unidad. alias es un parámetro nuevo
    (27/09/2026), al final y con valor por defecto."""
    norm = _normalizar_para_comparar(ubicacion_limpia)
    if norm in _UBICACIONES_PATIO_NORM:
        return True
    if alias:
        return norm in _PUNTOS_RESGUARDO_NORM.get(str(alias).strip().upper(), set())
    return False


# ── Lectura del archivo ───────────────────────────────────────────────────────

def leer_archivo(ruta):
    """
    Lee el 'Reporte de Estatus Actual' del GTRMax. Usa el parser 'lxml'
    porque el HTML de este reporte trae las etiquetas <tr> sin cerrar
    (le falta </tr> en cada fila salvo la primera); con 'html.parser'
    eso hace que todas las filas queden mal fusionadas en una sola.
    """
    contenido = None
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            with open(ruta, encoding=enc) as f:
                contenido = f.read()
            break
        except UnicodeDecodeError:
            continue
    if contenido is None:
        raise ValueError("No se pudo leer el archivo (problema de codificación).")

    soup = BeautifulSoup(contenido, "lxml")

    tabla_correcta = None
    encabezados = None
    for tabla in soup.find_all("table"):
        filas = tabla.find_all("tr")
        if not filas:
            continue
        primera = [c.get_text(strip=True) for c in filas[0].find_all(["th", "td"])]
        if "Alias" in primera and "Ubicación" in primera:
            tabla_correcta = tabla
            encabezados = primera
            break

    if tabla_correcta is None:
        raise ValueError(
            "No se encontró la tabla de datos (se esperaban columnas "
            "'Alias' y 'Ubicación'). ¿Es el archivo 'Reporte de Estatus Actual'?"
        )

    datos = []
    for fila in tabla_correcta.find_all("tr")[1:]:
        celdas = fila.find_all(["td", "th"])
        textos = [c.get_text(strip=True) for c in celdas]
        if len(textos) == len(encabezados):
            datos.append(textos)

    if not datos:
        raise ValueError("La tabla no tiene filas de datos.")

    df = pd.DataFrame(datos, columns=encabezados)

    # Hora de GENERACIÓN del reporte (nuevo — 20/09/2026), ej. de
    # "Reporte de Estatus Actual ( 20/09/2026 06:13:31 PM )" en el propio
    # HTML. Se guarda en df.attrs (no cambia el tipo de retorno ni la firma
    # pública de leer_archivo) para poder usarla como respaldo cuando el
    # export no trae la columna "Último Reporte" -- ver generar_reporte().
    df.attrs["fecha_generacion_reporte"] = _extraer_fecha_generacion_reporte(soup)

    return df


def _extraer_fecha_generacion_reporte(soup):
    """
    Busca en todo el HTML el texto "Reporte de Estatus Actual ( FECHA )"
    que trae el encabezado del export, y devuelve esa fecha como datetime.
    Retorna None si no se encuentra o no se puede parsear (no es un error
    fatal: generar_reporte() simplemente no podrá usar el modo de
    respaldo por "Fecha absoluta" si esto falla).
    """
    texto_completo = soup.get_text(" ", strip=True)
    m = re.search(r'Reporte de Estatus Actual\s*\(\s*([^)]+?)\s*\)', texto_completo)
    if not m:
        return None
    return _parsear_fecha_absoluta(m.group(1))


def encontrar_columna(columnas, candidatos):
    cols_lower = {c.strip().lower(): c for c in columnas}
    for cand in candidatos:
        if cand.strip().lower() in cols_lower:
            return cols_lower[cand.strip().lower()]
    for cand in candidatos:
        for orig_lower, orig in cols_lower.items():
            if cand.strip().lower() in orig_lower:
                return orig
    return None


# ── Parseo de horas ───────────────────────────────────────────────────────────

def parsear_hora_manual(texto: str) -> str:
    """
    Convierte la hora que escribe el usuario para el encabezado del mensaje
    a un formato FIJO de 12 horas con Am/Pm, ej. "08:00 Pm".

    Admite como entrada tanto 24h ("21:15", "08:00") como 12h con am/pm
    ("08:02 pm", "8:02 PM", "08:02 p.m."), pero la salida SIEMPRE se
    normaliza a 12h — antes se mostraba tal cual el usuario la escribía,
    lo que dejaba entradas en 24h ("20:00") sin am/pm en el mensaje final.
    """
    limpio = texto.strip()

    m = re.match(r'^(\d{1,2}):(\d{2})\s*([ap])\.?\s*m\.?$', limpio, re.IGNORECASE)
    if m:
        hora12 = int(m.group(1))
        minuto = int(m.group(2))
        if not (1 <= hora12 <= 12 and 0 <= minuto <= 59):
            raise ValueError(f"Formato de hora inválido: '{texto}'.")
        es_pm = m.group(3).lower() == "p"
    else:
        m2 = re.match(r'^(\d{1,2}):(\d{2})$', limpio)
        if not m2:
            raise ValueError(
                f"Formato de hora inválido: '{texto}'.\n"
                f"Usa por ejemplo: 21:15  o  08:02 pm"
            )
        hora24 = int(m2.group(1))
        minuto = int(m2.group(2))
        if not (0 <= hora24 <= 23 and 0 <= minuto <= 59):
            raise ValueError(f"Formato de hora inválido: '{texto}'.")
        es_pm = hora24 >= 12
        hora12 = hora24 % 12
        if hora12 == 0:
            hora12 = 12

    sufijo = "Pm" if es_pm else "Am"
    return f"{hora12:02d}:{minuto:02d} {sufijo}"


def _parsear_duracion_segundos(texto: str):
    """
    Convierte el texto de la columna 'Último Reporte' (tiempo transcurrido
    desde el último reporte, ej: '12s', '1m, 49s', '3m, 5s', '1h, 5m, 38s',
    '1d, 16m, 47s') a segundos totales. Retorna None si no se reconoce.
    """
    texto = str(texto).strip()
    if not texto:
        return None
    partes = re.findall(r'(\d+)\s*(d|h|m|s)', texto, re.IGNORECASE)
    if not partes:
        return None
    multiplicador = {"d": 86400, "h": 3600, "m": 60, "s": 1}
    return sum(int(valor) * multiplicador[unidad.lower()] for valor, unidad in partes)


def _parsear_fecha_absoluta(texto: str):
    """
    Parsea una fecha/hora absoluta del GTRMax, ej. '20/09/2026 06:14:17 PM'
    (o variantes sin segundos / en 24h). Retorna un datetime, o None si no
    se reconoce el formato.
    """
    texto = str(texto).strip()
    if not texto:
        return None
    for fmt in ("%d/%m/%Y %I:%M:%S %p", "%d/%m/%Y %H:%M:%S",
                "%d/%m/%Y %I:%M %p", "%d/%m/%Y %H:%M"):
        try:
            return datetime.strptime(texto, fmt)
        except ValueError:
            continue
    return None


def _segundos_desde_fecha_absoluta(fecha_fila_texto, fecha_generacion):
    """
    Modo de RESPALDO (nuevo — 20/09/2026) para cuando el export del GTRMax
    ya no trae la columna 'Último Reporte' (tiempo transcurrido en texto),
    sino una columna 'Fecha' con la marca de tiempo ABSOLUTA de cada
    registro (ej. '20/09/2026 06:14:17 PM'). En ese caso, el "tiempo
    transcurrido" se calcula como la diferencia entre la hora de
    GENERACIÓN del reporte completo (que trae el propio encabezado del
    HTML, ej. "Reporte de Estatus Actual ( 20/09/2026 06:13:31 PM )") y la
    Fecha de esa fila puntual.

    Si la Fecha de la fila es POSTERIOR a la hora de generación del reporte
    (puede pasar por unos segundos si el GTRMax tarda un poco en armar el
    export mientras siguen llegando transmisiones), se trata como 0
    segundos transcurridos (la unidad transmitió recién, prácticamente al
    momento del corte) en vez de dar un número negativo.

    Retorna None si la Fecha de la fila no se puede parsear.
    """
    fecha_fila = _parsear_fecha_absoluta(fecha_fila_texto)
    if fecha_fila is None:
        return None
    delta_segundos = (fecha_generacion - fecha_fila).total_seconds()
    return max(0, int(delta_segundos))


def limpiar_alias(raw: str) -> str:
    """Quita el prefijo 'UT:'/'Ut:' que trae el alias, igual que en el resto de ControlPlus."""
    return re.sub(r"(?i)^ut\s*:\s*", "", str(raw)).strip()


def _normalizar_para_exclusion(alias: str) -> str:
    """Forma canónica de un alias para comparar contra la lista de unidades a
    obviar (nuevo v3.9): 'r66'/'R-66'/'R-066' -> 'R-066'; '343'/'0343' -> '0343'."""
    a = re.sub(r"(?i)^ut\s*:\s*", "", str(alias)).strip().upper().replace(" ", "")
    m = re.fullmatch(r"R-?(\d{1,3})", a)
    if m:
        return f"R-{int(m.group(1)):03d}"
    m = re.fullmatch(r"\d{1,4}", a)
    if m:
        return a.zfill(4)
    return a


def _extraer_eje(matricula: str, alias_limpio: str):
    """
    Determina el Eje de la unidad a partir del sub-eje/etiqueta pegado
    al VIN en la columna 'Matrícula' (ej: 'VIN: 4919 Norte' → 'Norte').
    Aplica primero la excepción manual por unidad (UNIDADES_FORZADAS).
    Retorna el nombre del Eje, o None si la etiqueta no se reconoce.
    """
    if alias_limpio in UNIDADES_FORZADAS:
        return UNIDADES_FORZADAS[alias_limpio]

    m = re.match(r'^\s*(?:VIN|Vin|vin)\s*:\s*\S+\s+(.+?)\s*$', str(matricula).strip())
    if not m:
        return None
    etiqueta = re.sub(r'\s+', ' ', m.group(1)).strip().rstrip('.').lower()
    return SUBEJE_A_EJE.get(etiqueta)


def limpiar_ubicacion(texto: str) -> str:
    """
    Recorta la Ubicación para mostrar solo calle/sector, quitando
    Parroquia/Municipio/Estado/Venezuela del final.

    Las palabras administrativas solo se reconocen cuando arrancan su
    propio segmento (justo después de un punto, o al inicio del texto) —
    así no se recorta por error un lugar real que tenga alguna de esas
    palabras dentro de su propio nombre (ej. "Barrio Venezuela. Los
    Teques" no debe cortarse en "Venezuela", porque ahí es el nombre del
    barrio, no el país al final de la dirección).
    """
    texto = str(texto).strip()
    if not texto:
        return ""

    def patron_segmento(palabra):
        return r'(?:^|\.\s*)' + palabra + r'\b'

    m = re.search(patron_segmento("Parroquia"), texto, re.IGNORECASE)
    if m:
        antes = texto[:m.start()].strip()
        if antes:
            return antes.rstrip('.').strip() + '.'
        # No queda nada antes de "Parroquia" → usar el nombre de la parroquia
        resto = texto[m.end():].strip()
        m2 = re.match(r'\s*([^.]+)\.', resto)
        if m2:
            return m2.group(1).strip() + '.'
        return resto.rstrip('.').strip() + '.'

    # Sin "Parroquia": probar con Municipio / Estado / Venezuela como respaldo
    for palabra in ("Municipio", "Estado", "Venezuela"):
        m = re.search(patron_segmento(palabra), texto, re.IGNORECASE)
        if m:
            antes = texto[:m.start()].strip()
            if antes:
                return antes.rstrip('.').strip() + '.'

    return texto.rstrip('.').strip() + '.'


# ── Lógica principal ──────────────────────────────────────────────────────────

def generar_reporte(df: pd.DataFrame, hora_texto: str, log_fn=print, excluir_unidades=None) -> list:
    """
    Genera un reporte por eje de las UT en movimiento (menos de 3 minutos
    desde su último reporte), con su ubicación actual.

    excluir_unidades (nuevo v3.9): set opcional de alias (en cualquier
    formato reconocible, ej. {"R-066", "0343"}) que se omiten del reporte
    por pedido explícito del usuario, sin importar que cumplan la regla de
    "en movimiento" -- paso opcional que el bot pregunta antes de generar.

    Retorna: list[tuple]: [(eje, mensaje), ...]
    """
    excluir_norm = {_normalizar_para_exclusion(u) for u in excluir_unidades} if excluir_unidades else None
    hora_display = parsear_hora_manual(hora_texto)

    columnas = list(df.columns)
    col_alias = encontrar_columna(columnas, ["Alias"])
    col_matricula = encontrar_columna(columnas, ["Matrícula", "Matricula"])
    col_ubicacion = encontrar_columna(columnas, ["Ubicación", "Ubicacion"])
    col_ultimo = encontrar_columna(columnas, ["Último Reporte", "Ultimo Reporte"])
    col_fecha = encontrar_columna(columnas, ["Fecha"])
    fecha_generacion = df.attrs.get("fecha_generacion_reporte") if hasattr(df, "attrs") else None

    # Modo de respaldo (nuevo — 20/09/2026): el export del GTRMax dejó de
    # traer "Último Reporte" en algunos casos, mostrando en su lugar una
    # columna "Fecha" (marca de tiempo absoluta por unidad). Si pasa esto,
    # se calcula el tiempo transcurrido comparando esa Fecha contra la hora
    # de generación del reporte completo (ver _extraer_fecha_generacion_reporte
    # y _segundos_desde_fecha_absoluta) en vez de fallar.
    usar_fecha_absoluta = (col_ultimo is None and col_fecha is not None
                           and fecha_generacion is not None)

    faltantes = [n for n, c in [
        ("Alias", col_alias), ("Matrícula", col_matricula),
        ("Ubicación", col_ubicacion),
    ] if not c]
    if col_ultimo is None and not usar_fecha_absoluta:
        if col_fecha is None:
            faltantes.append("Último Reporte (tampoco se encontró una columna 'Fecha' de respaldo)")
        else:
            faltantes.append("Último Reporte (hay columna 'Fecha', pero no se pudo leer la hora de "
                              "generación del reporte del encabezado del archivo para usarla de respaldo)")
    if faltantes:
        raise ValueError(
            f"Columnas no encontradas: {', '.join(faltantes)}\n"
            f"Columnas detectadas: {', '.join(columnas)}"
        )

    if usar_fecha_absoluta:
        log_fn(
            "ℹ Este archivo no trae la columna 'Último Reporte' -- se calculó el "
            "tiempo transcurrido usando la columna 'Fecha' de cada unidad contra "
            f"la hora de generación del reporte ({fecha_generacion.strftime('%d/%m/%Y %I:%M:%S %p')})."
        )

    grupos = {eje: [] for eje in ORDEN_EJES}
    sin_clasificar = set()
    en_patio = 0
    excluidas_manual = 0

    for _, fila in df.iterrows():
        if usar_fecha_absoluta:
            segundos = _segundos_desde_fecha_absoluta(fila[col_fecha], fecha_generacion)
        else:
            segundos = _parsear_duracion_segundos(fila[col_ultimo])
        if segundos is None or segundos >= TOLERANCIA_SEGUNDOS:
            continue  # no está en movimiento (o no se pudo leer la duración)

        alias = limpiar_alias(fila[col_alias])
        if excluir_norm and _normalizar_para_exclusion(alias) in excluir_norm:
            excluidas_manual += 1
            continue  # unidad obviada por pedido explícito del usuario

        eje = _extraer_eje(fila[col_matricula], alias)
        if eje is None:
            sin_clasificar.add(str(fila[col_matricula]).strip())
            continue

        ubicacion = limpiar_ubicacion(fila[col_ubicacion])
        if es_ubicacion_patio(ubicacion, alias):
            en_patio += 1
            continue  # unidad resguardada en Patio de Vencedores, se omite

        grupos[eje].append(f"{alias} {ubicacion}".strip())

    if sin_clasificar:
        log_fn(
            f"⚠ {len(sin_clasificar)} etiqueta(s) de Matrícula sin clasificar "
            f"(no están en la tabla Eje): {', '.join(sorted(sin_clasificar))}"
        )
    if en_patio:
        log_fn(f"ℹ {en_patio} unidad(es) omitida(s) por estar en Patio o en su Punto de Resguardo.")
    if excluidas_manual:
        log_fn(f"ℹ {excluidas_manual} unidad(es) omitida(s) por pedido explícito del usuario.")

    resultados = []
    for eje in ORDEN_EJES:
        eje_display = eje.upper()
        distintivo = f"{'-' * 16}{eje_display}{'-' * 16}"
        unidades = grupos[eje]
        if unidades:
            lineas = [
                distintivo,
                "",
                f"El Centro de Control de Operaciones informa que para la "
                f"({hora_display}) se visualizan las siguientes UT en movimiento:"
            ]
            for i, texto_unidad in enumerate(unidades, 1):
                lineas.append(f"{i}. {texto_unidad}")
            mensaje = "\n".join(lineas)
        else:
            mensaje = (
                f"{distintivo}\n\n"
                f"El Centro de Control de Operaciones informa que para la "
                f"({hora_display}) no se visualizan UT en movimiento."
            )
        resultados.append((eje_display, mensaje))

    return resultados


def formatear_salida(resultados: list) -> str:
    bloques = []
    for nombre_eje, mensaje in resultados:
        bloques.append(f"{nombre_eje}\n\n{mensaje}")
    return "\n\n".join(bloques)


# ── Interfaz gráfica ──────────────────────────────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        root.title("Últimas UT en Movimiento")
        root.geometry("660x660")
        root.configure(bg="#0f1117")
        root.resizable(True, True)

        FONT_MONO = ("Courier New", 10)
        FONT_BOLD = ("Courier New", 11, "bold")
        BG     = "#0f1117"
        FG     = "#e2e8f0"
        ACCENT = "#3fb950"
        BTN_BG = "#1a7431"

        tk.Label(root, text="ÚLTIMAS UT EN MOVIMIENTO",
                 font=("Courier New", 14, "bold"),
                 bg=BG, fg=ACCENT).pack(pady=(20, 2))
        tk.Label(root, text="Reporte por eje según hora manual (GTR)",
                 font=FONT_MONO, bg=BG, fg="#64748b").pack(pady=(0, 14))

        # ── Hora manual ──
        frame_hora = tk.Frame(root, bg=BG)
        frame_hora.pack(pady=(0, 10))
        tk.Label(frame_hora, text="Hora actual (ej: 21:15):",
                 font=FONT_BOLD, bg=BG, fg=FG).pack(side="left", padx=(0, 8))
        self.entry_hora = tk.Entry(frame_hora, font=FONT_BOLD, width=14,
                                    bg="#1c2333", fg=ACCENT,
                                    insertbackground=FG, relief="flat")
        self.entry_hora.pack(side="left")

        # ── Selección de archivo ──
        tk.Button(root, text="📂  Seleccionar archivo GTR (.xls)",
                  font=FONT_BOLD, bg=BTN_BG, fg="white",
                  activebackground="#166534", activeforeground="white",
                  relief="flat", padx=16, pady=8,
                  command=self.cargar_archivo).pack(pady=(0, 6))

        self.lbl_archivo = tk.Label(root, text="Ningún archivo seleccionado",
                                    font=FONT_MONO, bg=BG, fg="#475569")
        self.lbl_archivo.pack(pady=(0, 10))

        tk.Button(root, text="⚡  Generar reporte",
                  font=("Courier New", 12, "bold"),
                  bg="#2f81f7", fg="white",
                  activebackground="#1c5fc4",
                  relief="flat", padx=20, pady=10,
                  command=self.generar).pack(pady=(0, 12))

        tk.Label(root, text="REPORTES GENERADOS:", font=FONT_BOLD,
                 bg=BG, fg="#94a3b8").pack(anchor="w", padx=20)

        self.txt = scrolledtext.ScrolledText(
            root, font=FONT_MONO, bg="#0d1117", fg="#86efac",
            insertbackground=FG, relief="flat", bd=0,
            padx=10, pady=10, wrap="word", height=18
        )
        self.txt.pack(fill="both", expand=True, padx=20, pady=(4, 12))

        btn_frame = tk.Frame(root, bg=BG)
        btn_frame.pack(pady=(0, 20))

        tk.Button(btn_frame, text="📋  Copiar todo",
                  font=FONT_BOLD, bg="#065f46", fg="white",
                  activebackground="#047857", relief="flat", padx=14, pady=7,
                  command=self.copiar_todo).pack(side="left", padx=6)

        tk.Button(btn_frame, text="💾  Guardar .txt",
                  font=FONT_BOLD, bg="#1e3a5f", fg="white",
                  activebackground="#1e40af", relief="flat", padx=14, pady=7,
                  command=self.guardar_txt).pack(side="left", padx=6)

        self.ruta_archivo = None
        self.resultados = []

    def cargar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo GTR",
            filetypes=[("Archivos GTR", "*.xls *.xlsx")]
        )
        if not ruta:
            return
        self.ruta_archivo = ruta
        self.lbl_archivo.config(text=os.path.basename(ruta), fg=("#3fb950"))

    def generar(self):
        if not self.ruta_archivo:
            messagebox.showinfo("Info", "Primero selecciona el archivo GTR.")
            return

        hora_texto = self.entry_hora.get().strip()
        if not hora_texto:
            messagebox.showinfo("Info", "Ingresa la hora manual (ej: 21:15).")
            return

        try:
            df = leer_archivo(self.ruta_archivo)
            self.resultados = generar_reporte(df, hora_texto)
            texto = formatear_salida(self.resultados)
            self.txt.delete("1.0", "end")
            self.txt.insert("end", texto)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def copiar_todo(self):
        if not self.resultados:
            messagebox.showinfo("Info", "No hay reportes generados.")
            return
        mensajes = [mensaje for _, mensaje in self.resultados]
        self.root.clipboard_clear()
        self.root.clipboard_append("\n\n".join(mensajes))
        messagebox.showinfo("Listo", f"¡{len(mensajes)} reporte(s) copiados al portapapeles!")

    def guardar_txt(self):
        if not self.resultados:
            messagebox.showinfo("Info", "No hay reportes generados.")
            return
        ruta = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Archivo de texto", "*.txt")],
            initialfile=f"ultimas_ut_movimiento_{datetime.now().strftime('%d%m%Y_%H%M')}.txt"
        )
        if not ruta:
            return
        with open(ruta, "w", encoding="utf-8") as f:
            f.write(formatear_salida(self.resultados))
        messagebox.showinfo("Guardado", f"Archivo guardado en:\n{ruta}")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
