"""
Data Alertas - Gestor de Incidencias de Transporte
Convierte reportes de texto en hojas Excel formateadas.
"""

import re
import os
import sys
import difflib
import unicodedata
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime, time
from openpyxl import Workbook, load_workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import AutoFilter

# ─── CONSTANTES DE NEGOCIO ────────────────────────────────────────────────────

COLUMNS = [
    "Responsable", "Número de reporte", "Fecha", "Hora",
    "Eje", "Sub‑eje", "Unidad", "Cts", "Tipo de reporte",
    "Incidencias", "Niveles", "Descripción", "Observación",
    "Respuestas", "Detalle"
]

VALID_INCIDENCIAS = {
    "exceso de velocidad": "Exceso de velocidad",
    "excesos de velocidad": "Exceso de velocidad",
    "fuera de ruta": "Fuera de ruta",
    "fueras de ruta": "Fuera de ruta",
    "ruta no atendida": "Ruta no atendida",
    "rutas no atendidas": "Ruta no atendida",
    "movimiento no reportado": "Movimiento no reportado",
    "movimientos no reportados": "Movimiento no reportado",
    "sin informacion": "Sin informacion",
    "sin información": "Sin informacion",
    "recorrido omitido": "Recorrido omitido",
    "recorridos omitidos": "Recorrido omitido",
    "falla en gps": "Falla en Gps",
    "paradas": "Paradas",
    "ubicación": "Ubicación",
    "ubicacion": "Ubicación",
    "desconexión de batería": "Desconexión de Batería",
    "desconexion de bateria": "Desconexión de Batería",
}

# Texto final que va en la columna "Incidencias" del Excel para estos tipos
# (abreviado, sin punto al final). Se aplica DESPUÉS de calcular Niveles y
# Descripción/Observación con el nombre canónico completo — esas reglas de
# negocio siguen dependiendo del nombre completo (ver _apply_business_rules
# y _compute_nivel), solo el texto que se muestra/guarda cambia al final.
DISPLAY_INCIDENCIAS = {
    "Exceso de velocidad": "Ex. de velocidad",
    "Movimiento no reportado": "Mov. NR",
    "Fuera de ruta": "Fuera de ruta",
    "Recorrido omitido": "Recorrido omitido",
    "Ruta no atendida": "Ruta no atendida",
}

EJE_MAP = {
    "am": "Am", "Am": "Am", "AM": "Am",
    "ocm": "Ocm", "Ocm": "Ocm", "OCM": "Ocm",
    "blv": "Blv", "Blv": "Blv", "BLV": "Blv",
    "pz": "Pz", "Pz": "Pz", "PZ": "Pz",
    "met": "Met", "Met": "Met", "MET": "Met",
    "serv": "Serv.", "serv.": "Serv.", "Serv.": "Serv.",
    "Serv": "Serv.", "SERV": "Serv.",
}

SUBEJE_MAP = {
    "norte": "Norte", "sur": "Sur", "centro": "Centro",
    "acevedo": "Acevedo", "brion": "Brión", "brión": "Brión",
    "buroz": "Buroz", "pedro gual": "Pedro Gual",
    # "Páez" a secas (nombre corto que suele venir en los reportes) también
    # se normaliza al nombre completo del sub-eje — nuevo 15/09/2026.
    "paez": "Páez Andrés Bello", "páez": "Páez Andrés Bello",
    "paez andres bello": "Páez Andrés Bello",
    "páez andrés bello": "Páez Andrés Bello",
    "páez andres bello": "Páez Andrés Bello",
    "paez andrés bello": "Páez Andrés Bello",
    "met": "Met",
    # "Sucre" no es un sub-eje válido en sí mismo — pertenece al eje/sub-eje Met.
    "sucre": "Met", "plaza sucre": "Met",
    "tuy i": "Tuy I", "tuy 1": "Tuy I",
    "tuy ii": "Tuy II", "tuy 2": "Tuy II",
    "plaza zamora": "Plaza Zamora",
    # Sub-eje del eje Serv.: "Esp." / "Cont." (15/09/2026)
    "esp": "Esp.", "esp.": "Esp.", "especial": "Esp.", "espc": "Esp.", "espc.": "Esp.",
    "cont": "Cont.", "cont.": "Cont.",
}

# Sub-eje por defecto cuando el Eje no trae sub-eje explícito en el texto
SUBEJE_DEFAULT = {
    "Met": "Met",
    "Pz": "Plaza Zamora",
}

# ─── RESPONSABLES CONOCIDOS (nuevo — 15 de septiembre de 2026) ────────────────
# Lista fija de responsables reales. Los reportes de WhatsApp casi nunca
# traen el nombre completo — traen un apodo o el nombre de perfil del
# celular (ej. "Lon", 'Jonaz "Bola 10"'), así que el nombre se reconoce por
# APROXIMACIÓN contra esta lista (ver _normalizar_responsable_conocido).
# Si el nombre no se parece lo suficiente a ninguno de estos, se guarda un
# guion "-" en vez del apodo tal cual (para no ensuciar la columna con
# nombres inventados o irreconocibles).
RESPONSABLES_CONOCIDOS = [
    "Maikel Moussa",
    "Leonardo Moreno",
    "Oriana Moreno",
    "Francys Mendoza",
    "Francelys Mendoza",
    "Jhona Ojeda",
    "Lonnell Tovar",
    "Nickolth Amarista",
    "Adriana Gudiño",
    "Natalia Ruza",
    "Wilmer Rodríguez",
    "María Echenique",
    "María Dominguez",
    "Hendry Valera",
]

# Umbral mínimo de similitud para aceptar una coincidencia, y margen mínimo
# que debe sacarle el mejor candidato al segundo mejor para no quedar en
# empate ambiguo (ej. "Moreno" solo, o "María" solo, calzan igual de bien
# con DOS responsables distintos — en ese caso se prefiere "-" a adivinar).
_RESP_UMBRAL_ACEPTAR = 0.55
_RESP_MARGEN_AMBIGUEDAD = 0.04


def _sin_acentos(s: str) -> str:
    """Quita tildes/diacríticos para comparar sin importar acentuación."""
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )


def _score_similitud(a: str, b: str) -> float:
    """
    Puntaje de similitud entre dos strings (ya en minúsculas, sin tildes):
      - 1.0 si son idénticos.
      - 0.85 si uno es prefijo del otro (con al menos 3 caracteres), ej.
        "lon" es el inicio de "lonnell" — esto cubre el caso muy común de
        un apodo de WhatsApp que es simplemente el nombre real cortado.
      - Si no, similitud de texto general (difflib.SequenceMatcher).
    """
    if a == b:
        return 1.0
    if len(a) >= 3 and len(b) >= 3 and (a.startswith(b) or b.startswith(a)):
        return 0.85
    return difflib.SequenceMatcher(None, a, b).ratio()


def _normalizar_responsable_conocido(nombre_raw: str) -> str:
    """
    Reconoce por APROXIMACIÓN el nombre/apodo de WhatsApp del responsable
    contra RESPONSABLES_CONOCIDOS. Compara el apodo (sin comillas, sin
    números, sin tildes, en minúsculas) contra el nombre completo y contra
    cada una de sus palabras (nombre y apellido) por separado, y se queda
    con la mejor coincidencia — siempre que supere el umbral mínimo Y le
    saque margen suficiente a la segunda mejor coincidencia distinta (para
    no adivinar entre dos responsables parecidos, ej. las dos "María" o
    los dos "Moreno" de la lista).

    Si no hay ninguna coincidencia suficientemente buena, retorna "-".
    """
    if not nombre_raw or not nombre_raw.strip():
        return "-"

    # Quitar apodos/alias entre comillas: 'Jonaz "Bola 10"' → 'Jonaz'
    limpio = re.sub(r'["“”\'‘’].*?["“”\'‘’]', " ", nombre_raw)
    limpio = re.sub(r"\d+", " ", limpio)              # quitar números sueltos
    limpio = re.sub(r"\s+", " ", limpio).strip()
    if not limpio:
        limpio = nombre_raw.strip()

    candidato = _sin_acentos(limpio).lower()

    puntajes = []   # [(nombre_completo, score), ...]
    for nombre_completo in RESPONSABLES_CONOCIDOS:
        objetivo = _sin_acentos(nombre_completo).lower()
        partes = objetivo.split()
        mejor = max(_score_similitud(candidato, p) for p in partes + [objetivo])
        puntajes.append((nombre_completo, mejor))

    puntajes.sort(key=lambda x: -x[1])
    mejor_nombre, mejor_score = puntajes[0]

    if mejor_score < _RESP_UMBRAL_ACEPTAR:
        return "-"

    # ¿Hay un segundo responsable DISTINTO casi tan bien puntuado? → ambiguo.
    for nombre, score in puntajes[1:]:
        if nombre != mejor_nombre and (mejor_score - score) < _RESP_MARGEN_AMBIGUEDAD:
            return "-"

    return mejor_nombre


def _limpiar_texto_whatsapp(texto: str) -> str:
    """
    Elimina marcas de formato Unicode INVISIBLES que WhatsApp suele insertar
    al copiar texto (ej. U+200E "left-to-right mark" pegada justo antes de
    cada "[hora, fecha]"). Bug real (15/09/2026): estas marcas rompen la
    detección de inicio de reporte (que exige que el timestamp esté al
    inicio de línea) para TODOS los mensajes salvo el primero, haciendo que
    el bot detecte un solo reporte enorme en vez de uno por mensaje.
    """
    invisibles = (
        "\u200b\u200c\u200d\u200e\u200f"   # zero-width, LRM, RLM
        "\u202a\u202b\u202c\u202d\u202e"   # embeddings/overrides bidi
        "\u2060\u2066\u2067\u2068\u2069"   # word joiner, isolates bidi
        "\ufeff"                           # BOM / zero-width no-break space
    )
    return texto.translate({ord(c): None for c in invisibles})

# Colores Excel
COLOR_HEADER_BG   = "1F3864"
COLOR_HEADER_FG   = "FFFFFF"
COLOR_ROW_ODD     = "FFFFFF"   # filas impares: sin relleno (blanco)
COLOR_ROW_EVEN    = "F2F2F2"   # filas pares: gris claro — "Filas con bandas"
COLOR_BORDER      = "BFBFBF"
COLOR_TEXT        = "000000"

COL_WIDTHS = {
    1: 14,   # Responsable
    2: 14,   # Número de reporte
    3: 14,   # Fecha
    4: 14,   # Hora
    5: 8,    # Eje
    6: 18,   # Sub-eje
    7: 10,   # Unidad
    8: 15,   # Cts
    9: 14,   # Tipo de reporte
    10: 22,  # Incidencias
    11: 10,  # Niveles
    12: 30,  # Descripción
    13: 45,  # Observación
    14: 12,  # Respuestas
    15: 18,  # Detalle
}

# ─── PARSEO DE TEXTO ──────────────────────────────────────────────────────────

# Patrón que identifica el inicio de un reporte individual:
# una línea que empieza con hora (ej: "9:05 a. m., 31/3/2026")
_REPORT_START = re.compile(
    r'^\s*\[?\s*\d{1,2}:\d{2}(?::\d{2})?\s*(?:a\.\s*m\.|p\.\s*m\.|am|pm|a\.m\.|p\.m\.)',
    re.IGNORECASE | re.MULTILINE
)


def parse_reports(raw_text):
    """
    Parseo en dos niveles:
      1. Divide el texto por secciones "Responsable:" — cada sección
         asigna ese nombre a TODOS los reportes que le siguen, hasta
         el próximo "Responsable:".
      2. Dentro de cada sección, detecta reportes individuales por su
         línea de timestamp (hora + fecha).

    Si una sección NO trae "Responsable:" explícito (formato WhatsApp
    crudo: "[hora, fecha] Nombre: mensaje"), el nombre se extrae del
    encabezado del primer bloque y se hereda a los bloques siguientes
    que no traigan su propio nombre (igual que WhatsApp, que solo
    muestra el nombre en el primer mensaje de una tanda consecutiva).
    Si dentro de una misma sección aparece un bloque con su propio
    encabezado de nombre distinto, ese nombre toma prioridad de ahí en
    adelante (para texto mezclado: etiquetado + WhatsApp crudo).
    """
    records = []

    # ── Sanitizar caracteres invisibles de WhatsApp (bug real, 15/09/2026) ───
    # Al copiar mensajes desde la app de WhatsApp (sobre todo grupos con
    # participantes de distintos idiomas/RTL), el texto suele traer marcas
    # de formato Unicode INVISIBLES (ej. U+200E "left-to-right mark") pegadas
    # justo antes de cada "[hora, fecha]". Como _REPORT_START exige que el
    # timestamp esté al INICIO de línea, esa marca invisible rompe la
    # detección de TODOS los mensajes salvo el primero, y el resto queda
    # fusionado en un solo bloque enorme del que solo se extrae UN reporte.
    # Se eliminan aquí, antes de dividir el texto en reportes individuales.
    raw_text = _limpiar_texto_whatsapp(raw_text)

    # ── Nivel 1: dividir por "Responsable:" ──────────────────────────────────
    # Cada elemento del split contiene: "Responsable: Nombre\n<reportes>"
    sections = re.split(r'(?=^\s*Responsable\s*:)', raw_text.strip(),
                        flags=re.IGNORECASE | re.MULTILINE)

    for section in sections:
        section = section.strip()
        if not section:
            continue

        current_responsable = ""

        # Extraer nombre del responsable de esta sección (si lo tiene)
        m_resp = re.match(r'Responsable\s*:\s*(.+)', section, re.IGNORECASE)
        if m_resp:
            current_responsable = _normalizar_responsable_conocido(m_resp.group(1).strip())
            # Quitar la línea "Responsable: X" del texto restante
            section = section[m_resp.end():].strip()

        # ── Nivel 2: dividir en reportes individuales por timestamp ──────────
        report_blocks = _split_into_report_blocks(section)

        for block in report_blocks:
            block = block.strip()
            if not block:
                continue

            # Si el propio bloque trae su encabezado WhatsApp ("[hora] Nombre:"),
            # ese nombre manda y se hereda hacia adelante; si no, se usa el
            # último nombre conocido (de "Responsable:" o de un bloque previo).
            candidato = _extraer_responsable_header(block)
            if candidato:
                current_responsable = candidato

            r = _parse_single_block(block, current_responsable)
            if r:
                records.append(r)

    return records


def _extraer_responsable_header(block):
    """
    Extrae el nombre del remitente del encabezado WhatsApp de un bloque:
    "[11:20 a. m., 11/8/2026] Jhona Alexander: Buenos Días" → nombre
    reconocido por aproximación contra RESPONSABLES_CONOCIDOS.
    Retorna "" si no hay nombre (ej. mensaje sin encabezado propio, heredado
    del anterior) o si lo que sigue a "]" es solo un saludo.

    El apodo de WhatsApp puede traer comillas y números (ej. el nombre de
    perfil 'Jonaz "Bola 10"') — se prioriza el corte en ": >" (el mensaje
    citado casi siempre empieza con "> "), que tolera cualquier caracter en
    el nombre; si no aparece ese patrón, se usa un corte más simple que solo
    admite letras/espacios/puntos.
    """
    primera_linea = block.splitlines()[0] if block.splitlines() else ""

    m = re.search(r'\]\s*(.+?)\s*:\s*>', primera_linea)
    if not m:
        m = re.search(
            r'\]\s*([A-Za-zÁÉÍÓÚáéíóúÑñ][A-Za-zÁÉÍÓÚáéíóúÑñ\s.]*?)\s*:',
            primera_linea
        )
    if not m:
        return ""
    candidato = m.group(1).strip()
    if re.match(r'^buen[ao]s?\b', candidato, re.IGNORECASE):
        return ""
    return _normalizar_responsable_conocido(candidato)


def _split_into_report_blocks(text):
    """
    Divide un bloque de texto en reportes individuales.
    Cada reporte comienza con una línea de timestamp.
    Si el texto no tiene timestamps, lo trata como un solo reporte.
    """
    # Encontrar todas las posiciones donde empieza un timestamp
    starts = [m.start() for m in _REPORT_START.finditer(text)]

    if not starts:
        # Sin timestamps reconocibles → devolver el texto completo como un bloque
        return [text]

    blocks = []
    for i, start in enumerate(starts):
        end = starts[i + 1] if i + 1 < len(starts) else len(text)
        blocks.append(text[start:end])

    return blocks


def _parse_single_block(text, responsable=""):
    """
    Extrae todos los campos de un reporte individual.
    El responsable se pasa como parámetro (heredado de la sección), pero si
    viene vacío se intenta extraer del encabezado de WhatsApp del propio
    bloque: "[11:20 a. m., 11/8/2026] Jhona Alexander: Buenos Días".

    Reconoce dos estilos de campo:
      - Con etiqueta y dos puntos:  "Eje: Blv"      "Unidad: 0656"
      - Con viñeta, sin dos puntos: "* Eje Blv - Sub Eje Acevedo"
                                     "* Unidad 0656 - Vin 0723"
    """
    rec = {c: "" for c in COLUMNS}
    rec["Respuestas"] = "No"   # ← valor por defecto corregido
    rec["Detalle"] = ""
    rec["Responsable"] = responsable

    # Se limpian los asteriscos de énfasis (Markdown de Telegram/negrita) que
    # traen los mensajes generados por otras herramientas del bot (ver el
    # reporte "Validación Parada" más abajo, que llega con *Eje*, *Unidad*,
    # etc.) — no aportan información, solo formato, y si no se quitan rompen
    # la extracción de Eje/Sub-eje/Unidad/CTS (las regex no contemplan '*').
    text = re.sub(r'\*', '', text)

    # ── Responsable (fallback: encabezado WhatsApp de la primera línea) ──────
    if not rec["Responsable"]:
        rec["Responsable"] = _extraer_responsable_header(text)

    # ── Hora ─────────────────────────────────────────────────────────────────
    m = re.search(
        r'\[?\s*(\d{1,2}:\d{2}(?::\d{2})?)\s*(a\.\s*m\.|p\.\s*m\.|AM|PM|a\.m\.|p\.m\.)',
        text, re.IGNORECASE
    )
    if m:
        rec["Hora"] = _normalize_time(m.group(1), m.group(2))
    else:
        m2 = re.search(r'\[?\s*(\d{1,2}:\d{2}(?::\d{2})?)\s*(am|pm)', text, re.IGNORECASE)
        if m2:
            rec["Hora"] = _normalize_time(m2.group(1), m2.group(2))

    # ── Fecha ─────────────────────────────────────────────────────────────────
    m = re.search(r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})', text)
    if m:
        rec["Fecha"] = _normalize_date(m.group(0))

    # ── Eje ───────────────────────────────────────────────────────────────────
    m = re.search(r'\bEje\s*:\s*(.+)', text, re.IGNORECASE)
    if m:
        rec["Eje"] = _normalize_eje(m.group(1).strip().rstrip(',;'))
    else:
        # Formato viñeta sin dos puntos: "* Eje Blv - Sub Eje Acevedo"
        m = re.search(r'\bEje\s+([A-Za-zÁÉÍÓÚáéíóúÑñ.]+)', text, re.IGNORECASE)
        if m:
            rec["Eje"] = _normalize_eje(m.group(1).strip().rstrip(',;.'))

    # ── Sub-eje ───────────────────────────────────────────────────────────────
    m = re.search(r'Sub[\-\s]?eje\s*:\s*(.+)', text, re.IGNORECASE)
    if m:
        rec["Sub‑eje"] = _normalize_subeje(m.group(1).strip().rstrip(',;'), rec["Eje"])
    else:
        # Formato viñeta sin dos puntos: "Sub Eje Acevedo" (resto de la línea)
        m = re.search(r'Sub[\-\s]?[Ee]je\s+([^\n]+)', text, re.IGNORECASE)
        if m:
            rec["Sub‑eje"] = _normalize_subeje(m.group(1).strip().rstrip(',;.'), rec["Eje"])
        else:
            rec["Sub‑eje"] = ""

    # Sub-eje por defecto cuando el Eje no trae sub-eje explícito
    if not rec["Sub‑eje"]:
        rec["Sub‑eje"] = SUBEJE_DEFAULT.get(rec["Eje"], "")

    # ── Unidad ────────────────────────────────────────────────────────────────
    m = re.search(r'Unidad\s*:\s*(.+)', text, re.IGNORECASE)
    if m:
        rec["Unidad"] = m.group(1).strip().rstrip(']').strip()
    else:
        # Formato viñeta sin dos puntos: "* Unidad 0656 - Vin 0723"
        m = re.search(r'\bUnidad\s+(\S+)', text, re.IGNORECASE)
        rec["Unidad"] = m.group(1).strip().rstrip(',;.') if m else ""

    # ── CTS (conductor) — extraído automáticamente del texto ─────────────────
    # Acepta: "Cts:Ronald Z.", "CTS: Yordy H.", "cts : Juan P."
    m = re.search(r'\bCts\s*:\s*(.+)', text, re.IGNORECASE)
    if m:
        rec["Cts"] = m.group(1).strip().rstrip(']').strip()
    else:
        rec["Cts"] = ""

    # ── Reporte especial "Validación Parada" (nuevo — 15 de septiembre de 2026)
    # Estos mensajes los genera Ojo de Halcón (Herramienta 9) y NO usan
    # "Incidencia:" — traen literalmente "Validación Parada" seguido de
    # "Ruta:" y "Ubicación de Parada:". Se procesan aparte porque no siguen
    # el patrón de las demás incidencias (Descripción/Observación/Niveles).
    if re.search(r'Validaci[oó]n\s*Parada', text, re.IGNORECASE):
        rec["Tipo de reporte"] = "Validacion"
        rec["Incidencias"]     = "Parada"
        rec["Niveles"]         = "-"

        m_ruta = re.search(r'Ruta\s*:\s*(.+)', text, re.IGNORECASE)
        rec["Descripción"] = m_ruta.group(1).strip().rstrip(']').strip() if m_ruta else ""

        # La ubicación de parada puede venir en la misma línea de la
        # etiqueta o en la línea siguiente; se toma todo hasta la pregunta
        # final ("¿Nos podrían indicar...?") o hasta el final del bloque.
        m_ubic = re.search(
            r'Ubicaci[oó]n\s+de\s+Parada\s*:\s*(.+?)(?:\n\s*¿|\Z)',
            text, re.IGNORECASE | re.DOTALL
        )
        rec["Observación"] = re.sub(r'\s+', ' ', m_ubic.group(1)).strip() if m_ubic else ""

        return rec

    # ── Tipo de reporte ───────────────────────────────────────────────────────
    rec["Tipo de reporte"] = "Incidencia"

    # ── Incidencia ────────────────────────────────────────────────────────────
    m = re.search(r'Incidencia\s*:\s*(.+)', text, re.IGNORECASE)
    raw_inc = m.group(1).strip().rstrip(']').strip() if m else ""
    rec["Incidencias"] = _normalize_incidencia(raw_inc)

    # ── Ruta ──────────────────────────────────────────────────────────────────
    m = re.search(r'Ruta\s*:\s*(.+)', text, re.IGNORECASE)
    ruta = m.group(1).strip().rstrip(']').strip() if m else ""

    # ── Observación ───────────────────────────────────────────────────────────
    m = re.search(r'Observaci[oó]n\s*:\s*(.+)', text, re.IGNORECASE)
    obs_texto = m.group(1).strip().rstrip(']').strip() if m else ""

    # Fallback específico para Exceso de velocidad: si no hay "Observación:"
    # explícita, usar el texto de "Incumplimiento:" (ej: "93 km/h en autopista"),
    # que es donde realmente viene la velocidad en el formato de viñetas.
    if not obs_texto and "exceso de velocidad" in rec["Incidencias"].lower():
        m_incump = re.search(r'Incumplimiento\s*:?\s*(.+)', text, re.IGNORECASE)
        if m_incump:
            obs_texto = m_incump.group(1).strip().rstrip(']').strip()

    # ── Descripción y Observación (reglas de negocio) ─────────────────────────
    rec["Descripción"], rec["Observación"] = _apply_business_rules(
        rec["Incidencias"], ruta, obs_texto
    )

    # ── Niveles ───────────────────────────────────────────────────────────────
    rec["Niveles"] = _compute_nivel(rec["Incidencias"], rec["Descripción"])

    # ── Texto final de Incidencias (abreviado, sin punto) ────────────────────
    # Se hace al final: Descripción/Observación y Niveles ya se calcularon
    # arriba usando el nombre canónico completo (ej. "Exceso de velocidad"),
    # que es lo que necesitan _apply_business_rules() y _compute_nivel().
    rec["Incidencias"] = formato_final_incidencia(rec["Incidencias"])

    return rec


def _normalize_time(time_str, ampm_str):
    """Convierte hora a formato hh:mm AM/PM (12h, dos dígitos, sin segundos)."""
    # OJO: el texto copiado de WhatsApp suele usar espacios "duros" (U+00A0)
    # o "angostos" (U+202F) en "a. m." / "p. m.", no el espacio normal
    # (U+0020). Un simple .replace(' ', '') no los elimina y hace que
    # TODO se detecte como PM por error. \s en regex sí cubre ambos.
    ampm_clean = re.sub(r'[.\s]', '', ampm_str, flags=re.UNICODE).upper()
    parts = time_str.split(':')
    h = int(parts[0])
    m = int(parts[1]) if len(parts) > 1 else 0
    # Normalizar a rango 12h (1-12), por si el texto trae 0 o >12
    h = h % 12
    if h == 0:
        h = 12
    suffix = "AM" if ampm_clean == "AM" else "PM"
    return f"{h:02d}:{m:02d} {suffix}"


def _normalize_date(date_str):
    """Convierte fecha a DD/MM/YYYY."""
    parts = re.split(r'[/\-]', date_str)
    if len(parts) == 3:
        d, mo, y = parts
        if len(y) == 2:
            y = "20" + y
        return f"{int(d):02d}/{int(mo):02d}/{y}"
    return date_str


def _normalize_eje(raw):
    """Normaliza el campo Eje a los valores estándar."""
    key = raw.strip().lower().rstrip('.')
    # Intentar coincidencia directa con mapa
    for k, v in EJE_MAP.items():
        if k.lower() == key or k.lower() == key + '.':
            return v
    # Retornar original si no hay coincidencia
    return raw


def _normalize_subeje(raw, eje):
    """Normaliza el campo Sub-eje."""
    key = raw.strip().lower().rstrip('.')
    for k, v in SUBEJE_MAP.items():
        if k == key:
            return v
    # Si Eje es Pz y viene "Plaza sucre" u otro, respetarlo
    return raw


def _normalize_incidencia(raw):
    """Normaliza el tipo de incidencia al valor canónico (nombre completo,
    usado internamente por las reglas de negocio y el cálculo de Niveles)."""
    key = raw.strip().rstrip('.').strip().lower()
    if key in VALID_INCIDENCIAS:
        return VALID_INCIDENCIAS[key]
    # Búsqueda parcial
    for k, v in VALID_INCIDENCIAS.items():
        if k in key or key in k:
            return v
    return raw.strip().rstrip('.').strip()


def formato_final_incidencia(canonica: str) -> str:
    """
    Texto final que se guarda en la columna 'Incidencias' del Excel: aplica
    la abreviación pedida para los 5 tipos de DISPLAY_INCIDENCIAS y, para
    cualquier otro tipo, deja el nombre canónico tal cual — en ambos casos
    sin punto final.
    """
    texto = canonica.strip().rstrip('.').strip()
    return DISPLAY_INCIDENCIAS.get(texto, texto)


def _apply_business_rules(incidencia, ruta, obs_texto):
    """
    Aplica las reglas de negocio para Descripción y Observación
    según el tipo de incidencia.
    """
    inc = incidencia.lower()

    if "exceso de velocidad" in inc:
        # Descripción: velocidad exacta (ej: "104 km/h")
        speed = _extract_speed(obs_texto)
        desc = speed if speed else obs_texto
        # Observación: Autopista / Carretera / Avenida (del texto o por defecto)
        obs = _extract_via_type(obs_texto)
        return desc, obs

    elif "fuera de ruta" in inc:
        desc = ruta
        # Observación: "Con recorrido hacia: X" o "Se visualiza en: X"
        if obs_texto:
            if re.search(r'con recorrido|hacia', obs_texto, re.IGNORECASE):
                obs = obs_texto
            elif re.search(r'visualiza|se encuentra', obs_texto, re.IGNORECASE):
                obs = obs_texto
            else:
                obs = f"Con recorrido hacia {obs_texto}"
        else:
            obs = ""
        return desc, obs

    elif "ruta no atendida" in inc:
        desc = ruta
        if re.search(r'sin inicio', obs_texto, re.IGNORECASE):
            obs = "Sin inicio de operaciones"
        else:
            obs = "Ruta no atendida"
        return desc, obs

    elif "recorrido omitido" in inc:
        desc = ruta
        if obs_texto:
            if re.search(r'sin recorrido', obs_texto, re.IGNORECASE):
                obs = obs_texto
            else:
                obs = f"Sin recorrido hacia {obs_texto}"
        else:
            obs = ""
        return desc, obs

    elif "movimiento no reportado" in inc:
        if re.search(r'reserva', obs_texto, re.IGNORECASE):
            desc = "Reserva en movimiento"
        else:
            desc = "Falla en movimiento"
        obs = obs_texto
        return desc, obs

    elif "sin informacion" in inc or "sin información" in inc:
        desc = obs_texto
        obs = obs_texto
        return desc, obs

    elif "falla en gps" in inc:
        if re.search(r'salto', obs_texto, re.IGNORECASE):
            desc = "Salto de ubicación"
            obs = obs_texto
        elif re.search(r'recorrido', obs_texto, re.IGNORECASE):
            desc = "Recorrido"
            obs = "Sin registro de recorrido"
        else:
            desc = "Desconexión"
            obs = obs_texto
        return desc, obs

    elif "paradas" in inc:
        desc = obs_texto
        obs = obs_texto
        return desc, obs

    elif "ubicación" in inc or "ubicacion" in inc:
        desc = obs_texto
        obs = obs_texto
        return desc, obs

    elif "desconexión de batería" in inc or "desconexion de bateria" in inc:
        desc = obs_texto
        obs = obs_texto
        return desc, obs

    return obs_texto, obs_texto


def _extract_speed(text):
    """Extrae valor de velocidad como '104 km/h' del texto."""
    m = re.search(r'(\d+)\s*km/h', text, re.IGNORECASE)
    if m:
        return f"{m.group(1)} km/h"
    return ""


def _extract_via_type(text):
    """Detecta tipo de vía: Autopista, Carretera o Avenida."""
    t = text.lower()
    if "autopista" in t:
        return "Autopista"
    if "carretera" in t:
        return "Carretera"
    if "avenida" in t or "av." in t:
        return "Avenida"
    return "Autopista"  # default según datos de referencia


def _compute_nivel(incidencia, descripcion):
    """Calcula el nivel (1/2/3) para Exceso de velocidad, según la velocidad:
       87–90 km/h → 1 | 91–100 km/h → 2 | 101+ km/h → 3"""
    if "exceso de velocidad" not in incidencia.lower():
        return "-"
    m = re.search(r'(\d+)', descripcion)
    if not m:
        return "-"
    speed = int(m.group(1))
    if 87 <= speed <= 90:
        return "1"
    elif 91 <= speed <= 100:
        return "2"
    elif speed >= 101:
        return "3"
    return "-"


# ─── EXCEL ENGINE ─────────────────────────────────────────────────────────────

def _time_sort_key(hora_str):
    """
    Convierte la Hora a una clave ordenable, tolerante a variaciones de
    formato (con/sin segundos, con/sin cero inicial, am/pm en minúsculas,
    con o sin puntos, formato 24h, o incluso un objeto datetime.time real
    si viene de una celda de Excel ya tipada).

    Devuelve (0, time(...)) cuando se pudo interpretar, o (1, time(23,59,59))
    cuando NO se pudo — así un valor raro se va al FINAL de la lista en
    vez de al principio (evita que un formato inesperado desordene todo
    lo demás, que es justo lo que pasaba antes).
    """
    if isinstance(hora_str, time):
        return (0, hora_str)
    if isinstance(hora_str, datetime):
        return (0, hora_str.time())
    if not hora_str:
        return (1, time(23, 59, 59))

    texto = str(hora_str).strip()

    # Formato 12h con AM/PM: "9:21:00 AM", "9:21 a.m.", "09:21:00 p. m."
    m = re.match(r'^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp])\.?\s*[Mm]\.?$', texto)
    if m:
        h = int(m.group(1)) % 12
        mi = int(m.group(2))
        s = int(m.group(3) or 0)
        if m.group(4).upper() == "P":
            h += 12
        try:
            return (0, time(h, mi, s))
        except ValueError:
            pass

    # Formato 24h: "21:15:00" o "21:15"
    m2 = re.match(r'^(\d{1,2}):(\d{2})(?::(\d{2}))?$', texto)
    if m2:
        try:
            return (0, time(int(m2.group(1)) % 24, int(m2.group(2)), int(m2.group(3) or 0)))
        except ValueError:
            pass

    return (1, time(23, 59, 59))


def _date_sort_key(fecha_str):
    """Convierte 'DD/MM/AAAA' a clave ordenable (año, mes, día).
       Fechas no interpretables se van al final."""
    if not fecha_str:
        return (9999, 12, 31)
    try:
        d, mo, y = re.split(r'[/\-]', str(fecha_str).strip())
        return (int(y), int(mo), int(d))
    except Exception:
        return (9999, 12, 31)


def load_existing_data(filepath):
    """Carga datos existentes del Excel y retorna lista de filas + último número."""
    rows = []
    last_num = 0
    try:
        wb = load_workbook(filepath)
        ws = wb["Data Alertas"]
        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
            if any(cell is not None for cell in row):
                rows.append(list(row))
                # Extraer número de reporte
                num_cell = row[1]
                if num_cell:
                    m = re.search(r'\d+', str(num_cell))
                    if m:
                        last_num = max(last_num, int(m.group()))
    except Exception:
        pass
    return rows, last_num


def build_excel(new_records, existing_filepath=None, output_filepath=None):
    """
    Construye el Excel con todos los registros (existentes + nuevos).
    Retorna la ruta del archivo guardado.
    """
    # Cargar datos existentes
    existing_rows, last_num = load_existing_data(existing_filepath) if existing_filepath else ([], 0)

    # Convertir nuevos registros a filas, asignando numeración
    all_rows = list(existing_rows)
    for rec in new_records:
        last_num += 1
        row = [
            rec.get("Responsable", ""),
            f"N°{last_num:03d}",
            rec.get("Fecha", ""),
            rec.get("Hora", ""),
            rec.get("Eje", ""),
            rec.get("Sub‑eje", "") or None,
            rec.get("Unidad", ""),
            rec.get("Cts", "") or None,
            rec.get("Tipo de reporte", "Incidencia"),
            rec.get("Incidencias", ""),
            rec.get("Niveles", "-"),
            rec.get("Descripción", ""),
            rec.get("Observación", ""),
            rec.get("Respuestas", "NO"),
            rec.get("Detalle", "") or None,
        ]
        all_rows.append(row)

    # Ordenar por Fecha y luego por Hora (columnas índice 2 y 3)
    all_rows.sort(key=lambda r: (_date_sort_key(r[2]), _time_sort_key(r[3])))

    # Re-numerar después de ordenar
    for i, row in enumerate(all_rows):
        row[1] = f"N°{i+1:03d}"

    # Crear workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Data Alertas"

    # ── Estilos ──────────────────────────────────────────────────────────────

    hdr_fill = PatternFill("solid", fgColor=COLOR_HEADER_BG)
    odd_fill  = PatternFill("solid", fgColor=COLOR_ROW_ODD)
    even_fill = PatternFill("solid", fgColor=COLOR_ROW_EVEN)

    hdr_font  = Font(bold=False, italic=False, color=COLOR_HEADER_FG, name="Calibri", size=11)
    data_font = Font(bold=False, italic=False, color=COLOR_TEXT, name="Calibri", size=11)

    center    = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_side = Side(style="thin", color=COLOR_BORDER)
    thin_bdr  = Border(left=thin_side, right=thin_side,
                       top=thin_side, bottom=thin_side)

    # ── Encabezados ──────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 20
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = center
        cell.border    = thin_bdr

    # ── Autofiltro ───────────────────────────────────────────────────────────
    last_col_letter = get_column_letter(len(COLUMNS))
    ws.auto_filter.ref = f"A1:{last_col_letter}1"

    # ── Filas de datos ───────────────────────────────────────────────────────
    for row_idx, row_data in enumerate(all_rows, start=2):
        # Filas con bandas: alterna sin relleno (blanco) / gris claro
        data_row = row_idx - 1   # 1-based dentro de los datos
        fill = odd_fill if data_row % 2 == 1 else even_fill

        ws.row_dimensions[row_idx].height = 15

        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.fill      = fill
            cell.font      = data_font
            cell.alignment = center
            cell.border    = thin_bdr

    # ── Anchos de columna ────────────────────────────────────────────────────
    for col_idx, width in COL_WIDTHS.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # ── Guardar ──────────────────────────────────────────────────────────────
    if not output_filepath:
        output_filepath = "Data_Alertas_output.xlsx"

    wb.save(output_filepath)
    return output_filepath


# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║          FUNCIÓN 2 – PROCESADOR DE REPORTES MASIVOS                        ║
# ║  Reescrita para el nuevo formato de entrada (v4)                           ║
# ╚══════════════════════════════════════════════════════════════════════════════╝

class AnalizadorReportes:
    """
    Parsea reportes de transporte con el formato:
      [HH:MM am/pm, DD/M/YYYY] Responsable: NombreResponsable:
      - Eje Blv - Sub eje Brión
      - Unidad R-009 - Vin 7225
      - Cts: Jean.
      Incidencia: ...
      Incumplimiento: ...
      Ruta: ...          (opcional)
      Observacion: ...   (opcional)

    Salida normalizada con campos separados y formato estándar.
    Solo usa re y librería estándar, sin dependencias externas.
    """

    # ── Mapas de normalización ────────────────────────────────────────────────

    # Eje: variantes de entrada → valor canónico
    _EJE_MAP = {
        "am": "Am",  "a.m": "Am",
        "blv": "Blv", "boulevard": "Blv",
        "met": "Met", "metro": "Met",
        "ocm": "Ocm",
        "pz": "Pz",  "plaza": "Pz",
        "serv": "Serv", "serv.": "Serv", "servicio": "Serv",
    }

    # Sub-eje válido por eje (clave en minúsculas para comparación)
    _SUBEJE_MAP = {
        "Am":   {"norte": "Norte", "sur": "Sur", "centro": "Centro"},
        "Blv":  {
            "acevedo": "Acevedo",
            "brion": "Brión", "brión": "Brión",
            "buroz": "Buroz",
            "paez": "Páez", "páez": "Páez",
            "pedro gual": "Pedro Gual", "pedro": "Pedro Gual",
        },
        "Met":  {"metro": "Metro", "sucre": "Sucre"},
        "Ocm":  {
            "tuy i": "Tuy I", "tuy 1": "Tuy I", "tuyi": "Tuy I",
            "tuy ii": "Tuy II", "tuy 2": "Tuy II", "tuyii": "Tuy II",
        },
        "Pz":   {"plaza zamora": "Plaza Zamora", "zamora": "Plaza Zamora"},
        "Serv": {
            "esp": "Serv. Esp", "esp.": "Serv. Esp", "especial": "Serv. Esp",
            "serv. esp": "Serv. Esp", "serv esp": "Serv. Esp",
            "cont": "Serv. Cont", "cont.": "Serv. Cont", "continuo": "Serv. Cont",
            "serv. cont": "Serv. Cont", "serv cont": "Serv. Cont",
        },
    }

    # Sub-eje por defecto cuando el eje no trae sub-eje explícito
    _SUBEJE_DEFAULT = {"Met": "Met", "Pz": "Plaza Zamora"}

    # Palabras que indican una ubicación en el texto de incumplimiento
    _UBICACION_KEYWORDS = re.compile(
        r'\b(autopista|carretera|avenida|av\.|calle|urb\.|urbanizaci[oó]n|'
        r'sector|barrio|highway|boulevard|blvd)\b',
        re.IGNORECASE
    )

    # ── Normalización de Hora ─────────────────────────────────────────────────

    def _normalizar_hora(self, texto):
        """
        Extrae hora de la primera línea del bloque y la normaliza a:
        "HH:MM am" / "HH:MM pm"  (con cero a la izquierda, sin puntos)
        Acepta: "3:55 p. m.", "9:05 a.m.", "9:05am", "09:05 PM", etc.
        """
        m = re.search(
            r'(\d{1,2}:\d{2})(?::\d{2})?\s*'
            r'(a\.?\s*m\.?|p\.?\s*m\.?|am|pm)',
            texto, re.IGNORECASE
        )
        if not m:
            return None
        hm = m.group(1)                          # "3:55" o "9:05"
        suf_raw = re.sub(r'[.\s]', '', m.group(2), flags=re.UNICODE).lower()
        sufijo = "am" if suf_raw.startswith('a') else "pm"
        h, mi = hm.split(':')
        return f"{int(h):02d}:{mi} {sufijo}"     # "03:55 pm", "09:05 am"

    # ── Normalización de Fecha ────────────────────────────────────────────────

    def _normalizar_fecha(self, texto):
        """
        Extrae fecha y la normaliza a DD/MM/YYYY con ceros a la izquierda.
        Acepta: "5/4/2026", "05/04/2026", "5-4-2026".
        """
        m = re.search(r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})', texto)
        if not m:
            return None
        d, mo, y = int(m.group(1)), int(m.group(2)), m.group(3)
        if len(y) == 2:
            y = "20" + y
        return f"{d:02d}/{mo:02d}/{y}"            # "05/04/2026"

    # ── Normalización de Responsable ──────────────────────────────────────────

    def _normalizar_responsable(self, nombre_raw):
        """
        Capitaliza correctamente el nombre del responsable.
        "WILMER" → "Wilmer" | "wilmer garcia" → "Wilmer Garcia"
        """
        return " ".join(p.capitalize() for p in nombre_raw.strip().split())

    # ── Parseo de Eje y Sub-eje (misma línea) ─────────────────────────────────

    def _parsear_eje_subeje(self, bloque):
        """
        Busca la línea que contiene Eje y Sub eje juntos, separados por guión.
        Formato esperado: "- Eje Blv - Sub eje Brión"
        También acepta: "Eje: Blv - Sub eje: Brión"

        Retorna (eje_normalizado, subeje_normalizado) o (None, None).
        """
        for linea in bloque.splitlines():
            # Detectar que la línea menciona el eje
            m_eje = re.search(r'\beje\s*:?\s*([A-Za-zÁÉÍÓÚáéíóú]+)', linea, re.IGNORECASE)
            if not m_eje:
                continue

            eje_raw = m_eje.group(1).strip().lower().rstrip('.')
            eje = self._EJE_MAP.get(eje_raw)
            if not eje:
                continue

            # Buscar sub-eje en la misma línea después de un guión
            # Acepta: "Sub eje Brión" / "Sub-eje: Brión" / "sub eje Brión"
            m_sub = re.search(
                r'sub[\s\-]?eje\s*:?\s*([^\-\n\[]+)',
                linea, re.IGNORECASE
            )
            if m_sub:
                sub_raw = m_sub.group(1).strip().rstrip('.-').lower()
                opciones = self._SUBEJE_MAP.get(eje, {})
                subeje = None
                # Coincidencia exacta primero
                if sub_raw in opciones:
                    subeje = opciones[sub_raw]
                else:
                    # Coincidencia parcial
                    for k, v in opciones.items():
                        if k in sub_raw or sub_raw in k:
                            subeje = v
                            break
                if not subeje:
                    subeje = self._SUBEJE_DEFAULT.get(eje)
            else:
                subeje = self._SUBEJE_DEFAULT.get(eje)

            return eje, subeje

        return None, None

    # ── Normalización de Unidad ───────────────────────────────────────────────

    def _normalizar_unidad(self, bloque):
        """
        Extrae la unidad de una línea que puede tener o no los dos puntos.
        Patrones: "Unidad R-009", "Unidad: R-009", "Unidad 0485"
        Normaliza: R009 → R-009 | R-9 → R-009 (zero-pad a 3 dígitos)
        Ignora el campo VIN que puede venir en la misma línea tras " - ".
        """
        for linea in bloque.splitlines():
            m = re.search(r'\bunidad\s*:?\s*([R\-0-9]+)', linea, re.IGNORECASE)
            if not m:
                continue
            raw = m.group(1).strip()
            # Normalizar formato R: R009 o R9 → R-009
            r_match = re.match(r'^[Rr]-?(\d+)$', raw)
            if r_match:
                num = int(r_match.group(1))
                return f"R-{num:03d}"
            # Número puro (0485)
            return raw
        return None

    # ── Normalización de Cts ──────────────────────────────────────────────────

    def _normalizar_cts(self, bloque):
        """
        Extrae el nombre del conductor tal como está, con capitalización correcta.
        Conserva el punto al final si es una inicial (ej: "Jean." → "Jean.")
        """
        m = re.search(r'\bcts\s*:\s*([^\n\-\[]+)', bloque, re.IGNORECASE)
        if not m:
            return None
        raw = m.group(1).strip().rstrip(']').strip()
        if not raw:
            return None
        # Capitalizar cada palabra, preservar punto al final
        tiene_punto = raw.endswith('.')
        partes = raw.rstrip('.').split()
        resultado = " ".join(p.capitalize() for p in partes)
        return resultado + ('.' if tiene_punto else '')

    # ── Normalización de Incidencia ───────────────────────────────────────────

    def _normalizar_incidencia(self, bloque):
        """
        Extrae el campo 'Incidencia:' (no 'Incumplimiento:').
        Capitaliza. Agrega punto final si no lo tiene.
        """
        m = re.search(
            r'^[\s\-]*incidencia\s*:?\s*([^\n]+)',
            bloque, re.IGNORECASE | re.MULTILINE
        )
        if not m:
            return None
        raw = m.group(1).strip().rstrip(']').strip()
        if not raw:
            return None
        texto = raw[0].upper() + raw[1:]
        if not texto.endswith('.'):
            texto += '.'
        return texto

    # ── Normalización de Incumplimiento ──────────────────────────────────────

    def _normalizar_incumplimiento(self, bloque):
        """
        Extrae el campo 'Incumplimiento:'.
        Si contiene una ubicación (autopista, calle, etc.) la separa
        y la retorna como observación adicional.

        Retorna (incumplimiento_texto, observacion_extra) donde cualquiera
        puede ser None.
        """
        m = re.search(
            r'^[\s\-]*incumplimiento\s*:?\s*([^\n]+)',
            bloque, re.IGNORECASE | re.MULTILINE
        )
        if not m:
            return None, None

        raw = m.group(1).strip().rstrip(']').strip()
        if not raw:
            return None, None

        # Intentar separar la parte numérica/descriptiva de la ubicación
        # Ejemplo: "97km/h en autopista" → incump="97km/h" obs="Autopista"
        obs_extra = None
        incump_texto = raw

        # Buscar si hay una preposición + ubicación al final
        m_sep = re.search(
            r'^(.+?)\s+(?:en|en la|por|sobre)\s+(.+)$',
            raw, re.IGNORECASE
        )
        if m_sep:
            parte_dato     = m_sep.group(1).strip()   # "97km/h"
            parte_ubicacion = m_sep.group(2).strip()  # "autopista"
            if self._UBICACION_KEYWORDS.search(parte_ubicacion):
                incump_texto = parte_dato
                obs_extra = parte_ubicacion[0].upper() + parte_ubicacion[1:]
        else:
            # Sin preposición: revisar si la cadena entera es una ubicación
            if self._UBICACION_KEYWORDS.search(raw):
                # Todo es ubicación → mover a observación
                obs_extra = raw[0].upper() + raw[1:]
                incump_texto = None

        return incump_texto, obs_extra

    # ── Normalización de Ruta ─────────────────────────────────────────────────

    def _normalizar_ruta(self, bloque):
        """Extrae Ruta si está presente. Retorna None si no hay."""
        m = re.search(r'^\s*ruta\s*:\s*([^\n]+)', bloque, re.IGNORECASE | re.MULTILINE)
        if not m:
            return None
        raw = m.group(1).strip().rstrip(']').strip()
        if not raw:
            return None
        return raw[0].upper() + raw[1:]

    # ── Normalización de Observación ──────────────────────────────────────────

    def _normalizar_observacion(self, bloque):
        """Extrae Observacion si está presente. Retorna None si no hay."""
        m = re.search(
            r'^\s*observaci[oó]n\s*:\s*([^\n]+)',
            bloque, re.IGNORECASE | re.MULTILINE
        )
        if not m:
            return None
        raw = m.group(1).strip().rstrip(']').strip()
        if not raw:
            return None
        return raw[0].upper() + raw[1:]

    # ── Separador de reportes ─────────────────────────────────────────────────

    def separar_reportes(self, texto):
        """
        Divide el texto en bloques individuales.
        El nuevo formato tiene el responsable en la PRIMERA línea de cada reporte:
          "[3:55 p. m., 5/4/2026] wilmer:"
        Si un bloque no trae responsable → hereda el último detectado.

        Retorna lista de (responsable: str, bloque_completo: str).
        """
        # Patrón de inicio de reporte: línea con corchete o timestamp + nombre opcional
        # Detecta: "[3:55 p. m., 5/4/2026] wilmer:" o "3:55 p. m., 5/4/2026"
        patron_inicio = re.compile(
            r'^\s*\[?\s*\d{1,2}:\d{2}(?::\d{2})?\s*'
            r'(?:a\.?\s*m\.?|p\.?\s*m\.?|am|pm)',
            re.IGNORECASE | re.MULTILINE
        )

        inicios = [m.start() for m in patron_inicio.finditer(texto)]
        if not inicios:
            return []

        bloques_raw = []
        for i, inicio in enumerate(inicios):
            fin = inicios[i + 1] if i + 1 < len(inicios) else len(texto)
            bloques_raw.append(texto[inicio:fin].strip())

        resultado = []
        ultimo_responsable = ""

        for bloque in bloques_raw:
            if not bloque:
                continue
            primera_linea = bloque.splitlines()[0]

            # Formato esperado: "[HH:MM am/pm, DD/M/YYYY] NombreResponsable: [saludo]"
            # El nombre viene DESPUÉS del corchete de cierre "]" y ANTES de ":"
            # Puede haber saludo después del ":" en la misma línea → ignorarlo
            m_resp = re.search(
                r'\]\s*([A-Za-zÁÉÍÓÚáéíóúÑñ][A-Za-zÁÉÍÓÚáéíóúÑñ\s]*)\s*:',
                primera_linea
            )
            if m_resp:
                candidato = m_resp.group(1).strip()
                # Ignorar si es un saludo completo ("Buenos días", etc.)
                if not re.match(r'^buen[ao]s?\b', candidato, re.IGNORECASE):
                    ultimo_responsable = self._normalizar_responsable(candidato)

            resultado.append((ultimo_responsable, bloque))

        return resultado

    # ── Formateador de un reporte ─────────────────────────────────────────────

    def _formatear_reporte(self, responsable, bloque):
        """
        Extrae todos los campos del bloque y genera el texto de salida
        con el formato estándar definido.
        """
        lineas = []

        # ── 1. Responsable ────────────────────────────────────────────────────
        lineas.append(f"Responsable: {responsable}")

        # ── 2. Fecha (DD/MM/YYYY con ceros) ───────────────────────────────────
        fecha = self._normalizar_fecha(bloque)
        if fecha:
            lineas.append(f"Fecha: {fecha}")

        # ── 3. Hora (HH:MM am/pm sin puntos) ──────────────────────────────────
        hora = self._normalizar_hora(bloque)
        if hora:
            lineas.append(f"Hora: {hora}")

        # ── 4. Eje y Sub-eje ──────────────────────────────────────────────────
        eje, subeje = self._parsear_eje_subeje(bloque)
        if eje:
            lineas.append(f"Eje: {eje}")
        if subeje:
            lineas.append(f"Sub-eje: {subeje}")

        # ── 5. Unidad ─────────────────────────────────────────────────────────
        unidad = self._normalizar_unidad(bloque)
        if unidad:
            lineas.append(f"Unidad: {unidad}")

        # ── 6. Cts ────────────────────────────────────────────────────────────
        cts = self._normalizar_cts(bloque)
        if cts:
            lineas.append(f"Cts: {cts}")

        # ── 7. Incidencia ─────────────────────────────────────────────────────
        incidencia = self._normalizar_incidencia(bloque)
        if incidencia:
            lineas.append(f"Incidencia: {incidencia}")

        # ── 8. Incumplimiento + posible Observación derivada ──────────────────
        incump, obs_de_incump = self._normalizar_incumplimiento(bloque)
        if incump:
            lineas.append(f"Incumplimiento: {incump}")

        # ── 9. Ruta (opcional) ────────────────────────────────────────────────
        ruta = self._normalizar_ruta(bloque)
        if ruta:
            lineas.append(f"Ruta: {ruta}")

        # ── 10. Observación: primero la del texto, luego la derivada de incump ─
        obs_explicita = self._normalizar_observacion(bloque)
        obs_final = obs_explicita or obs_de_incump
        if obs_final:
            lineas.append(f"Observacion: {obs_final}")

        return "\n".join(lineas)

    # ── Orquestador principal ─────────────────────────────────────────────────

    def procesar_texto_masivo(self, texto):
        """
        Recibe texto crudo con N reportes.
        Retorna (texto_normalizado: str, total_procesados: int).
        """
        if not texto.strip():
            return "", 0

        bloques = self.separar_reportes(texto)
        if not bloques:
            return "", 0

        reportes_formateados = []
        for responsable, bloque in bloques:
            formateado = self._formatear_reporte(responsable, bloque)
            if formateado.strip():
                reportes_formateados.append(formateado)

        return "\n\n".join(reportes_formateados), len(reportes_formateados)


# ─── INTERFAZ GRÁFICA (Tkinter) ───────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Data Alertas – Gestor de Incidencias")
        self.geometry("960x740")
        self.resizable(True, True)
        self.configure(bg="#F0F4FA")
        self._existing_path = tk.StringVar()
        self._output_path   = tk.StringVar(value="Data_Alertas_output.xlsx")
        self._build_ui()

    def _build_ui(self):
        # ── Encabezado global ─────────────────────────────────────────────────
        header = tk.Frame(self, bg="#1F3864", height=50)
        header.pack(fill="x")
        tk.Label(
            header,
            text="  📋  Data Alertas – Gestor de Incidencias de Transporte",
            bg="#1F3864", fg="white", font=("Arial", 13, "bold"), anchor="w"
        ).pack(side="left", pady=10, padx=10)

        # ── Notebook con 2 pestañas ───────────────────────────────────────────
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Custom.TNotebook",
            background="#F0F4FA", borderwidth=0
        )
        style.configure(
            "Custom.TNotebook.Tab",
            background="#D9E1F2", foreground="#1F3864",
            font=("Arial", 10, "bold"),
            padding=[14, 6]
        )
        style.map(
            "Custom.TNotebook.Tab",
            background=[("selected", "#1F3864")],
            foreground=[("selected", "#FFFFFF")]
        )

        self._notebook = ttk.Notebook(self, style="Custom.TNotebook")
        self._notebook.pack(fill="both", expand=True, padx=0, pady=0)

        # ── Pestaña 1: Generador de Excel ─────────────────────────────────────
        tab1 = tk.Frame(self._notebook, bg="#F0F4FA")
        self._notebook.add(tab1, text="  📊  Función 1 – Generar Excel  ")
        self._build_tab1(tab1)

        # ── Pestaña 2: Procesador de Reportes Masivos ─────────────────────────
        tab2 = tk.Frame(self._notebook, bg="#F0F4FA")
        self._notebook.add(tab2, text="  ⚡  Función 2 – Reportes Masivos  ")
        self._build_tab2(tab2)

        # ── Barra de estado global ────────────────────────────────────────────
        self._status = tk.StringVar(value="Listo.")
        tk.Label(
            self, textvariable=self._status,
            bg="#D9E1F2", font=("Arial", 9),
            anchor="w", relief="sunken"
        ).pack(fill="x", side="bottom", ipady=3)

    # ══════════════════════════════════════════════════════════════════════════
    # PESTAÑA 1 — Generador de Excel (código 100% original, sin modificar)
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab1(self, parent):
        body = tk.Frame(parent, bg="#F0F4FA")
        body.pack(fill="both", expand=True, padx=15, pady=10)

        # Archivos
        file_frame = tk.LabelFrame(
            body, text=" Archivos ", bg="#F0F4FA",
            font=("Arial", 10, "bold"), fg="#1F3864"
        )
        file_frame.pack(fill="x", pady=(0, 8))

        tk.Label(file_frame, text="Excel existente (opcional):",
                 bg="#F0F4FA", font=("Arial", 9)).grid(
            row=0, column=0, sticky="w", padx=8, pady=4)
        tk.Entry(file_frame, textvariable=self._existing_path,
                 width=55, font=("Arial", 9)).grid(
            row=0, column=1, padx=5, pady=4)
        tk.Button(
            file_frame, text="Examinar…", command=self._browse_existing,
            bg="#1F3864", fg="white", font=("Arial", 9), cursor="hand2"
        ).grid(row=0, column=2, padx=5, pady=4)

        tk.Label(file_frame, text="Guardar como:",
                 bg="#F0F4FA", font=("Arial", 9)).grid(
            row=1, column=0, sticky="w", padx=8, pady=4)
        tk.Entry(file_frame, textvariable=self._output_path,
                 width=55, font=("Arial", 9)).grid(
            row=1, column=1, padx=5, pady=4)
        tk.Button(
            file_frame, text="Guardar en…", command=self._browse_output,
            bg="#1F3864", fg="white", font=("Arial", 9), cursor="hand2"
        ).grid(row=1, column=2, padx=5, pady=4)

        # Campos adicionales
        extra_frame = tk.LabelFrame(
            body, text=" Campos adicionales (opcionales) ",
            bg="#F0F4FA", font=("Arial", 10, "bold"), fg="#1F3864"
        )
        extra_frame.pack(fill="x", pady=(0, 8))

        tk.Label(extra_frame, text="Responsable (sobreescribe):",
                 bg="#F0F4FA", font=("Arial", 9)).grid(
            row=0, column=0, sticky="w", padx=8, pady=4)
        self._responsable_var = tk.StringVar()
        tk.Entry(extra_frame, textvariable=self._responsable_var,
                 width=22, font=("Arial", 9)).grid(
            row=0, column=1, sticky="w", padx=5, pady=4)

        tk.Label(extra_frame, text="Cts (conductor):",
                 bg="#F0F4FA", font=("Arial", 9)).grid(
            row=0, column=2, sticky="w", padx=12, pady=4)
        self._cts_var = tk.StringVar()
        tk.Entry(extra_frame, textvariable=self._cts_var,
                 width=22, font=("Arial", 9)).grid(
            row=0, column=3, sticky="w", padx=5, pady=4)

        tk.Label(extra_frame, text="Tipo de reporte:",
                 bg="#F0F4FA", font=("Arial", 9)).grid(
            row=0, column=4, sticky="w", padx=12, pady=4)
        self._tipo_var = tk.StringVar(value="Incidencia")
        ttk.Combobox(
            extra_frame, textvariable=self._tipo_var,
            values=["Incidencia", "Validacion"], width=14,
            font=("Arial", 9), state="readonly"
        ).grid(row=0, column=5, sticky="w", padx=5, pady=4)

        # Texto de entrada
        tk.Label(
            body, text="Pegue aquí el texto con los reportes de incidencias:",
            bg="#F0F4FA", font=("Arial", 10, "bold"), fg="#1F3864", anchor="w"
        ).pack(fill="x", pady=(0, 3))

        self._text_input = scrolledtext.ScrolledText(
            body, height=18, font=("Consolas", 9),
            wrap="word", relief="solid", bd=1
        )
        self._text_input.pack(fill="both", expand=True, pady=(0, 8))
        self._text_input.insert("1.0", _EXAMPLE_TEXT)

        # Botones
        btn_frame = tk.Frame(body, bg="#F0F4FA")
        btn_frame.pack(fill="x")

        tk.Button(
            btn_frame, text="🗑  Limpiar texto", command=self._clear_text,
            bg="#E8EDF5", fg="#1F3864", font=("Arial", 10),
            cursor="hand2", relief="flat", padx=12, pady=6
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            btn_frame, text="👁  Vista previa", command=self._preview,
            bg="#4472C4", fg="white", font=("Arial", 10, "bold"),
            cursor="hand2", relief="flat", padx=14, pady=6
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            btn_frame, text="✅  Generar Excel", command=self._generate,
            bg="#1F7840", fg="white", font=("Arial", 11, "bold"),
            cursor="hand2", relief="flat", padx=18, pady=6
        ).pack(side="right")

    # ══════════════════════════════════════════════════════════════════════════
    # PESTAÑA 2 — Procesador de Reportes Masivos (NUEVO)
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab2(self, parent):
        """Construye la UI de la Función 2: procesador masivo de reportes."""
        self._analizador = AnalizadorReportes()

        body = tk.Frame(parent, bg="#F0F4FA")
        body.pack(fill="both", expand=True, padx=15, pady=10)

        # ── Instrucción ───────────────────────────────────────────────────────
        tk.Label(
            body,
            text="Pegue el texto crudo (WhatsApp u otra fuente) y presione ⚡ Procesar Reportes",
            bg="#F0F4FA", fg="#1F3864", font=("Arial", 10, "bold"), anchor="w"
        ).pack(fill="x", pady=(0, 5))

        # ── Panel dividido izquierda / derecha ────────────────────────────────
        paned = tk.PanedWindow(body, orient="horizontal",
                               bg="#D9E1F2", sashwidth=6, sashrelief="flat")
        paned.pack(fill="both", expand=True)

        # ── Panel IZQUIERDO: texto de entrada ─────────────────────────────────
        left_frame = tk.Frame(paned, bg="#F0F4FA")
        paned.add(left_frame, minsize=300)

        tk.Label(
            left_frame, text="📥  Texto de entrada (crudo):",
            bg="#F0F4FA", fg="#1F3864", font=("Arial", 9, "bold"), anchor="w"
        ).pack(fill="x", padx=4, pady=(0, 2))

        self._f2_input = scrolledtext.ScrolledText(
            left_frame, font=("Consolas", 9),
            wrap="word", relief="solid", bd=1, bg="#FAFCFF"
        )
        self._f2_input.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        # Texto de ejemplo para Función 2
        self._f2_input.insert("1.0", _EXAMPLE_TEXT_F2)

        # ── Botón central PROCESAR ────────────────────────────────────────────
        center_bar = tk.Frame(body, bg="#F0F4FA", height=46)
        center_bar.pack(fill="x", pady=5)
        center_bar.pack_propagate(False)

        self._f2_contador = tk.StringVar(value="")
        tk.Label(
            center_bar, textvariable=self._f2_contador,
            bg="#F0F4FA", fg="#1F7840", font=("Arial", 10, "bold")
        ).pack(side="left", padx=8)

        tk.Button(
            center_bar, text="📋  Copiar todo",
            command=self._f2_copiar,
            bg="#4472C4", fg="white", font=("Arial", 10, "bold"),
            cursor="hand2", relief="flat", padx=12, pady=5
        ).pack(side="right", padx=(0, 8))

        tk.Button(
            center_bar, text="💾  Guardar .txt",
            command=self._f2_guardar,
            bg="#1F7840", fg="white", font=("Arial", 10, "bold"),
            cursor="hand2", relief="flat", padx=12, pady=5
        ).pack(side="right", padx=(0, 4))

        tk.Button(
            center_bar, text="🗑  Limpiar",
            command=self._f2_limpiar,
            bg="#E8EDF5", fg="#1F3864", font=("Arial", 10),
            cursor="hand2", relief="flat", padx=10, pady=5
        ).pack(side="right", padx=(0, 4))

        tk.Button(
            center_bar, text="⚡  Procesar Reportes",
            command=self._f2_procesar,
            bg="#C55A11", fg="white", font=("Arial", 11, "bold"),
            cursor="hand2", relief="flat", padx=18, pady=5
        ).pack(side="right", padx=(0, 8))

        # ── Panel DERECHO: texto procesado ────────────────────────────────────
        right_frame = tk.Frame(paned, bg="#F0F4FA")
        paned.add(right_frame, minsize=300)

        tk.Label(
            right_frame, text="📤  Reportes normalizados (resultado):",
            bg="#F0F4FA", fg="#1F3864", font=("Arial", 9, "bold"), anchor="w"
        ).pack(fill="x", padx=4, pady=(0, 2))

        self._f2_output = scrolledtext.ScrolledText(
            right_frame, font=("Consolas", 9),
            wrap="word", relief="solid", bd=1,
            bg="#F0FFF4", state="disabled"
        )
        self._f2_output.pack(fill="both", expand=True, padx=4, pady=(0, 4))

    # ── Callbacks Pestaña 1 (idénticos a la versión anterior) ─────────────────

    def _browse_existing(self):
        p = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")]
        )
        if p:
            self._existing_path.set(p)

    def _browse_output(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if p:
            self._output_path.set(p)

    def _clear_text(self):
        self._text_input.delete("1.0", "end")

    def _get_records(self):
        raw = self._text_input.get("1.0", "end").strip()
        if not raw:
            return []
        records = parse_reports(raw)
        resp_override = self._responsable_var.get().strip()
        cts_override  = self._cts_var.get().strip()
        tipo_override = self._tipo_var.get().strip()
        for r in records:
            if resp_override:
                r["Responsable"] = resp_override
            if cts_override:
                r["Cts"] = cts_override
            if tipo_override:
                r["Tipo de reporte"] = tipo_override
        return records

    def _preview(self):
        records = self._get_records()
        if not records:
            messagebox.showinfo("Vista previa", "No se detectaron registros en el texto.")
            return
        win = tk.Toplevel(self)
        win.title(f"Vista previa – {len(records)} registro(s)")
        win.geometry("1050x400")
        cols = ["Responsable", "Fecha", "Hora", "Eje", "Sub‑eje",
                "Unidad", "Incidencias", "Niveles", "Descripción", "Observación"]
        tree = ttk.Treeview(win, columns=cols, show="headings")
        widths = [110, 90, 90, 55, 130, 75, 150, 70, 160, 250]
        for c, w in zip(cols, widths):
            tree.heading(c, text=c)
            tree.column(c, width=w, minwidth=40)
        for rec in records:
            tree.insert("", "end", values=[rec.get(c, "") for c in cols])
        sb = ttk.Scrollbar(win, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=sb.set)
        tree.pack(fill="both", expand=True)
        sb.pack(fill="x")

    def _generate(self):
        records = self._get_records()
        if not records:
            messagebox.showwarning("Sin datos", "No se detectaron reportes en el texto.")
            return
        existing = self._existing_path.get().strip() or None
        output   = self._output_path.get().strip() or "Data_Alertas_output.xlsx"
        try:
            self._status.set("⏳ Generando Excel…")
            self.update_idletasks()
            path = build_excel(records, existing, output)
            self._status.set(f"✅ Archivo guardado: {path}")
            messagebox.showinfo(
                "Éxito",
                f"Excel generado correctamente con {len(records)} registro(s) nuevo(s).\n\n📁 {path}"
            )
        except Exception as e:
            self._status.set(f"❌ Error: {e}")
            messagebox.showerror("Error", str(e))

    # ── Callbacks Pestaña 2 (NUEVO) ───────────────────────────────────────────

    def _f2_procesar(self):
        """Procesa el texto crudo de entrada y muestra los reportes normalizados."""
        texto = self._f2_input.get("1.0", "end").strip()
        if not texto:
            messagebox.showwarning(
                "Sin datos",
                "Pegue texto con reportes en el panel izquierdo."
            )
            return
        try:
            self._status.set("⏳ Procesando reportes…")
            self.update_idletasks()
            resultado, total = self._analizador.procesar_texto_masivo(texto)
            # Mostrar en el panel de salida
            self._f2_output.config(state="normal")
            self._f2_output.delete("1.0", "end")
            self._f2_output.insert("1.0", resultado)
            self._f2_output.config(state="disabled")
            # Actualizar contador
            self._f2_contador.set(f"✅  {total} reporte(s) procesado(s)")
            self._status.set(f"✅ {total} reporte(s) procesado(s) correctamente.")
        except Exception as e:
            self._status.set(f"❌ Error: {e}")
            messagebox.showerror("Error al procesar", str(e))

    def _f2_copiar(self):
        """Copia todo el contenido del panel de resultados al portapapeles."""
        self._f2_output.config(state="normal")
        contenido = self._f2_output.get("1.0", "end").strip()
        self._f2_output.config(state="disabled")
        if not contenido:
            messagebox.showinfo("Vacío", "No hay resultado que copiar. Procese primero.")
            return
        self.clipboard_clear()
        self.clipboard_append(contenido)
        self._status.set("📋 Resultado copiado al portapapeles.")

    def _f2_guardar(self):
        """Guarda el resultado como archivo .txt."""
        self._f2_output.config(state="normal")
        contenido = self._f2_output.get("1.0", "end").strip()
        self._f2_output.config(state="disabled")
        if not contenido:
            messagebox.showinfo("Vacío", "No hay resultado que guardar. Procese primero.")
            return
        ruta = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt"), ("Todos", "*.*")],
            title="Guardar reportes procesados"
        )
        if not ruta:
            return
        try:
            with open(ruta, "w", encoding="utf-8") as f:
                f.write(contenido)
            self._status.set(f"💾 Guardado: {ruta}")
            messagebox.showinfo("Guardado", f"Archivo guardado correctamente:\n{ruta}")
        except Exception as e:
            messagebox.showerror("Error al guardar", str(e))

    def _f2_limpiar(self):
        """Limpia ambos paneles de la Función 2."""
        self._f2_input.delete("1.0", "end")
        self._f2_output.config(state="normal")
        self._f2_output.delete("1.0", "end")
        self._f2_output.config(state="disabled")
        self._f2_contador.set("")
        self._status.set("Listo.")


# ─── TEXTOS DE EJEMPLO ────────────────────────────────────────────────────────

_EXAMPLE_TEXT = """\
Responsable: Adri

9:05 a. m., 31/3/2026
Buenos Dias
Eje: Am
Sub eje: Norte
Unidad: R-003
Cts:Ronald Z.
Incidencia: Fuera de ruta
Ruta: Cortada Del Guayabo - Cortada De Maturin
Observacion: Se visualiza en Autopista Valle - Coche y Calle Zea. Caracas.

9:14 a. m., 31/3/2026
Buenos Dias
Eje: Am
Sub eje: Norte
Unidad: R-022
CTS: Yordy H.
Incidencia: Fuera de ruta
Ruta: San Antonio - Potrero Gordo
Observacion: Con recorrido hacia Urbanizacion Los Salias.

Responsable: Natalia

9:30 a. m., 31/3/2026
Eje: Pz
Unidad: 0638
Cts: Julio R.
Incidencia: Exceso de velocidad
Ruta: Caracas - Guarenas
Observacion: va a 98 km/h autopista
"""

# Ejemplo para la Función 2 (texto más crudo / desordenado, como llega de WhatsApp)
_EXAMPLE_TEXT_F2 = """[3:55 p. m., 5/4/2026] Wilmer: Buenas Tardes
- Eje Blv - Sub eje Brión
- Unidad R-009 - Vin 7225
- Cts: Jean.
Incidencia: Exceso de velocidad
Incumplimiento: 97km/h en autopista

[9:05 a. m., 5/4/2026] Adri: Buenos Días
- Eje Am - Sub eje Norte
- Unidad R-003 - Vin 1234
- Cts: Ronald Z.
Incidencia: Fuera de ruta
Ruta: Cortada Del Guayabo - Cortada De Maturin
Observacion: Se visualiza en Autopista Valle - Coche y Calle Zea.

[10:30 a. m., 5/4/2026] Natalia: Buenos Días
- Eje Pz
- Unidad 0638 - Vin 5566
- Cts: Julio R.
Incidencia: Exceso de velocidad
Incumplimiento: 98km/h en autopista
Ruta: Caracas - Guarenas

[11:15 a. m., 5/4/2026] Natalia:
- Eje Ocm - Sub eje Tuy I
- Unidad R-040 - Vin 9901
- Cts: Jose R.
Incidencia: Movimiento no reportado
Incumplimiento: Falla en movimiento
"""


# ─── CLI (modo sin interfaz) ──────────────────────────────────────────────────

def cli_mode():
    """Modo de línea de comandos: lee texto de stdin o archivo."""
    import argparse
    parser = argparse.ArgumentParser(
        description="Data Alertas – Genera Excel de incidencias desde texto"
    )
    parser.add_argument("input", nargs="?", help="Archivo de texto con reportes (o stdin)")
    parser.add_argument("-e", "--existing", help="Excel existente para actualizar")
    parser.add_argument("-o", "--output", default="Data_Alertas_output.xlsx",
                        help="Ruta de salida del Excel")
    args = parser.parse_args()

    if args.input:
        with open(args.input, encoding="utf-8") as f:
            raw_text = f.read()
    else:
        print("Pegue el texto de reportes y presione Ctrl+D (Linux/Mac) o Ctrl+Z+Enter (Windows):")
        raw_text = sys.stdin.read()

    records = parse_reports(raw_text)
    if not records:
        print("⚠️  No se detectaron reportes en el texto.")
        sys.exit(1)

    path = build_excel(records, args.existing, args.output)
    print(f"✅ Excel generado: {path} ({len(records)} registro(s))")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if "--cli" in sys.argv or (len(sys.argv) > 1 and not sys.argv[1].startswith("-")):
        sys.argv = [a for a in sys.argv if a != "--cli"]
        cli_mode()
    else:
        app = App()
        app.mainloop()
