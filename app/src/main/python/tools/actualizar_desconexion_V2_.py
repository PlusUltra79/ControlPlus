# -*- coding: utf-8 -*-
"""
actualizar_desconexion.py
==========================
Herramienta 9 (independiente) del sistema ControlPlus — MoussaCorp / TransMiranda.

Descripción:
    Lee el PDF "Resumen de Unidades que No Han Reportado" del GTRMax y el Excel
    de "Desconexion" (formato Data Alertas). Actualiza las columnas
    "Ultima Transmision" y "Tiempo de Desconexion" del Excel a partir de la
    columna "Fecha reporte" del PDF, y agrega como filas nuevas las unidades
    que aparecen en el PDF pero no existían en el Excel.

    Reglas adicionales:
      - Toda unidad con MENOS de 24 horas exactas de desconexión se OMITE
        (se considera que aún no lleva un día completo desconectada).
      - Se incluyen TODAS las demás unidades, sin importar los meses/años
        que lleven desconectadas.
      - Las unidades listadas en UNIDADES_A_OBVIAR_DEL_PDF se ignoran por
        completo cada vez que se lee el PDF (no se actualizan ni se agregan).
      - El resultado se ordena alfabéticamente por Eje -> Sub Eje -> Unidad.

    Mantiene EXACTAMENTE el mismo formato del Excel original (columnas,
    encabezados, colores, bordes, tipo de letra, banda de color par/impar)
    porque edita el mismo archivo con openpyxl en vez de reconstruirlo.

Dependencias:
    pip install openpyxl pdfplumber --break-system-packages

Uso rápido (modo script, sin GUI):
    1. Ajusta FECHA_ACTUAL más abajo (o pásala como argumento de línea de comandos).
    2. python actualizar_desconexion.py ruta_pdf.pdf ruta_excel.xlsx [carpeta_salida] [DD/MM/AAAA]

Funciones públicas (mismo estilo que el resto de scripts de ControlPlus):
    leer_pdf_desconexion(ruta_pdf) -> list[dict]
    leer_excel_desconexion(ruta_excel) -> (Workbook, Worksheet)
    actualizar_desconexion(ruta_pdf, ruta_excel, fecha_actual, carpeta_salida, log_fn) -> str
    guardar_excel(wb, carpeta_salida) -> str
"""

import os
import re
import sys
import copy
import calendar
import unicodedata
import numbers
from datetime import datetime

import pdfplumber
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ─────────────────────────────────────────────────────────────────────────────
#  ⚙️  CONFIGURACIÓN — colocar aquí la fecha actual antes de correr el script
# ─────────────────────────────────────────────────────────────────────────────
# Formato: DD/MM/AAAA  (o DD/MM/AAAA HH:MM si además quieres fijar la hora).
# Representa el día en que se está corriendo la actualización.
# Si solo escribes la fecha (sin hora), el script asume que ese día ya
# transcurrió por completo (23:59:59) — esto es a propósito: evita el error
# de "falta 1 día" que ocurre si se asumiera medianoche (00:00:00), y hace
# que el resultado no dependa de a qué hora del día se corra el script.
# Si se deja en None, el script usa la fecha y hora del sistema (ahora mismo).
FECHA_ACTUAL = None   # Ejemplo: "29/07/2026"


# ─────────────────────────────────────────────────────────────────────────────
#  ⚙️  UNIDADES A OBVIAR DEL PDF (lista manual)
# ─────────────────────────────────────────────────────────────────────────────
# Unidades que se deben IGNORAR por completo cada vez que se lee el PDF: no se
# actualiza su 'Ultima Transmision' con el dato del PDF, y tampoco se agregan
# como fila nueva si no existieran en el Excel. Si la unidad ya existe en el
# Excel, su fila existente se mantiene intacta (no se toca).
#
# Usa el mismo formato que la columna 'Unidad' del Excel: '0089', 'R-004', etc.
# Para agregar o quitar una unidad de esta lista, solo edita la lista de abajo.
UNIDADES_A_OBVIAR_DEL_PDF = {
    "0584", "0420", "0101", "0455", "0444", "0547", "0421", "0541", "0333", "0544",
    "0517", "0560", "0714", "0469", "0677", "0464", "0580", "0692", "0441", "0724",
    "0555", "0432", "0135", "0630", "0440", "0478", "0567", "R-004", "0089", "0710",
    "0460", "0244", "R-037", "0537", "0639", "0585", "0394", "0550", "0653", "0586",
    "0190", "0311", "0483", "0437", "0476", "0474", "0524", "0682", "0691", "0569",
    "0649", "0684", "0606", "0679", "0676", "0575", "R-002", "0635", "0551", "0513",
    "0552", "0701", "0449", "0471", "0622", "R-032", "0468", "0688", "0470", "0683",
    "0536", "0678", "0465", "0576",
}


# ─────────────────────────────────────────────────────────────────────────────
#  Mapa de sub-ejes → eje oficial
#  (según la lista entregada por MoussaCorp / TransMiranda)
# ─────────────────────────────────────────────────────────────────────────────
# clave: sub-eje normalizado (minúsculas, sin tildes, sin puntos)
# valor: (Eje mostrado en el Excel, Sub Eje mostrado en el Excel)
MAPA_SUBEJE = {
    "sur":            ("Am",    "Sur"),
    "norte":          ("Am",    "Norte"),
    "centro":         ("Am",    "Centro"),

    "acevedo":        ("Blv",   "Acevedo"),
    "brion":          ("Blv",   "Brión"),
    "buroz":          ("Blv",   "Buroz"),
    "paez":           ("Blv",   "Páez Andrés Bello"),
    "pedro gual":     ("Blv",   "Pedro Gual"),

    "sucre":          ("Met",   "Sucre"),
    "suc":            ("Met",   "Sucre"),
    "met":            ("Met",   "Met"),

    "tuy i":          ("Ocm",   "Tuy I"),
    "tuy ii":         ("Ocm",   "Tuy II"),

    "pz":             ("Pz",    "Plaza Zamora"),
    "plaza zamora":   ("Pz",    "Plaza Zamora"),

    "cttd":           ("Serv.", "Cont."),
    "cont":           ("Serv.", "Cont."),
    "esp":            ("Serv.", "Esp."),
    "ggm":            ("Serv.", "GGM"),
    "bus esc":        ("Serv.", "Bus. Esc."),
}

# Columnas del Excel (1-indexed) — deben coincidir con el archivo original
COL_RESPONSABLE = 1
COL_FECHA = 2
COL_CORTE = 3
COL_ULTIMA_TX = 4
COL_EJE = 5
COL_SUBEJE = 6
COL_UNIDAD = 7
COL_TIEMPO_DESC = 8
COL_COND_GPS = 9
COL_STATUS = 10

FILA_INICIAL_DATOS = 2

# Unidades sin chip GPS instalado: la columna "Condicion de GPS" siempre
# muestra "Sin Chip" para estas, y "Sin Transmitir" para el resto — es una
# regla fija por unidad, se aplica siempre (no depende del PDF ni de ningún
# Excel adicional). Para agregar/quitar una unidad de esta lista, editar
# el set de abajo (mismo formato que la columna 'Unidad': '0358', 'R-004').
UNIDADES_SIN_CHIP = {
    "0358", "0529", "0559", "0564", "0590", "0610", "0668", "0675",
}


def _condicion_gps(alias: str) -> str:
    """'Sin Chip' para las unidades de UNIDADES_SIN_CHIP, 'Sin Transmitir'
    para el resto."""
    return "Sin Chip" if str(alias).strip().upper() in UNIDADES_SIN_CHIP else "Sin Transmitir"

# Patrón para leer la columna "Unidad" del PDF, p.ej.:
#   "R-002(VIN: 4877 Sur)"       ->  alias=R-002   vin=4877  sub=Sur
#   "UT: 0089(Vin: 0892 Tuy I)"  ->  alias=UT: 0089 vin=0892  sub=Tuy I
#   "UT: 0190(VIN: 0906)"        ->  alias=UT: 0190 vin=0906  sub=None
PATRON_UNIDAD_PDF = re.compile(
    r'^(?P<alias>[A-Za-z]*[:\-]?\s*\d+)\*?\s*\(\s*[Vv][Ii][Nn]:\s*'
    r'(?P<vin>\d+)(?:\s+(?P<sub>[^)]*))?\)$'
)

# Patrón para leer la fecha "HCR:DD/MM/AAAA HH:MM:SS AM/PM"
PATRON_FECHA_PDF = re.compile(
    r'HCR:(?P<fecha>\d{2}/\d{2}/\d{4})\s+(?P<hora>\d{1,2}:\d{2}:\d{2}\s*[AP]M)',
    re.IGNORECASE
)


def _log_default(msg):
    print(msg)


# ─────────────────────────────────────────────────────────────────────────────
#  Excel de Disponibilidad (opcional) — Status/Estatus por unidad
# ─────────────────────────────────────────────────────────────────────────────
# Mismo archivo que usa "Excesos de Velocidad (PDF)" (una hoja por Eje/
# Sub-eje), pero aquí en vez de tomar el CTS se toma el Estatus/Status/Disponibilidad de
# cada unidad, para actualizar la columna "Status" del Excel de Desconexión.

_VALORES_VACIOS_STATUS = {"", "nan", "s.i", "s.i.", "s/n", "-", "none"}


def _celda_alias_valida_status(texto: str) -> bool:
    t = texto.strip().lower()
    return t in {"ut", "ut2", "unidad"} or "numeracion" in t


def _celda_status_valida(texto: str) -> bool:
    t = texto.strip().lower()
    # En la hoja de Ocm (Tuy I / Tuy II) la columna no se llama "Estatus" ni
    # "Status", sino "Disponibilidad" — mismo dato, solo cambia el nombre.
    return "estatus" in t or "status" in t or "disponibilidad" in t


def _formatear_alias_numerico(valor) -> str:
    """
    Corrige un bug real (16/09/2026, unidad 0140 de Ocm): cuando la columna
    de alias de una hoja de Disponibilidad es TOTALMENTE numérica (ninguna
    unidad tipo "R-XXX" que la obligue a quedar como texto), python-calamine
    la interpreta como columna numérica y PIERDE los ceros a la izquierda
    (ej. "0140" llega como el número 140.0). Esto es justo la fragilidad que
    ya se había detectado con datos sintéticos (ver Historial de
    Correcciones de Ojo de Halcón) — pero en la hoja real de Ocm sí llega a
    ocurrir, porque ahí todas las unidades son numéricas.

    Si el valor es un número entero (int/float/numpy.int64/numpy.float64
    sin parte decimal real), se reconstruye como texto de 4 dígitos con
    ceros a la izquierda — formato fijo de alias numérico usado en TODO
    el proyecto (ej. "0645", "0140"). Si ya viene como texto, se retorna
    tal cual.
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


class StatusDB(dict):
    """
    dict {ALIAS_UPPER: Status} normal en todo lo demás (get(), truthiness,
    etc. — bot.py y el resto del proyecto no necesitan cambiar), pero que
    además recuerda qué combinaciones (Eje, SubEje) correspondieron a una
    hoja de Disponibilidad completamente VACÍA (nuevo 16/09/2026, ver
    ejes_sin_reportar) — usado para escribir "Sin Reportar" en Status
    cuando no hay NINGUNA información de ese eje/sub-eje en el Excel de
    Disponibilidad adjunto.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.ejes_sin_reportar = set()   # {(Eje, SubEje), ...}


def _mapear_hoja_a_eje_subeje(nombre_hoja: str):
    """
    Mapea el nombre de una hoja de Disponibilidad (ej. "Serv Esp", "Pedro
    Gual", "Paez") al (Eje, SubEje) canónico del Excel de Desconexión,
    reutilizando MAPA_SUBEJE. A diferencia de mapear_eje_subeje() (que
    espera una sola etiqueta ya extraída del PDF), acá el nombre de la hoja
    puede tener varias palabras que no coinciden exactamente con ninguna
    llave (ej. "Serv Esp" vs. la llave "esp"), así que se prueban
    combinaciones de palabras consecutivas, de más larga a más corta.
    Retorna None si no se reconoce ninguna combinación.
    """
    clave = _quitar_tildes(nombre_hoja)
    if clave in MAPA_SUBEJE:
        return MAPA_SUBEJE[clave]
    palabras = clave.split()
    for tam in range(len(palabras), 0, -1):
        for ini in range(0, len(palabras) - tam + 1):
            sub = " ".join(palabras[ini:ini + tam])
            if sub in MAPA_SUBEJE:
                return MAPA_SUBEJE[sub]
    return None


def _detectar_columnas_status(fila_encabezado: list):
    idx_alias, idx_status = None, None
    for i, celda in enumerate(fila_encabezado):
        texto = str(celda) if celda is not None else ""
        if idx_alias is None and _celda_alias_valida_status(texto):
            idx_alias = i
        if idx_status is None and _celda_status_valida(texto):
            idx_status = i
    return idx_alias, idx_status


def _parsear_hoja_status(df_raw: "pd.DataFrame"):
    """Busca la fila de encabezado (alias + Estatus/Status/Disponibilidad) entre las
    primeras 3 filas de la hoja. Retorna {ALIAS_UPPER: status} o None si
    la hoja no tiene ese formato tabular (ej. hojas en texto libre)."""
    for fila_idx in range(min(3, len(df_raw))):
        fila = df_raw.iloc[fila_idx].tolist()
        idx_alias, idx_status = _detectar_columnas_status(fila)
        if idx_alias is None or idx_status is None:
            continue

        resultado = {}
        for r in range(fila_idx + 1, len(df_raw)):
            fila_datos = df_raw.iloc[r].tolist()
            if idx_alias >= len(fila_datos) or idx_status >= len(fila_datos):
                continue
            alias_raw = fila_datos[idx_alias]
            if alias_raw is None or not str(alias_raw).strip():
                continue
            alias = re.sub(r"(?i)^ut2?:?\s*", "", _formatear_alias_numerico(alias_raw)).strip().upper()
            status_raw = fila_datos[idx_status]
            status_txt = str(status_raw).strip() if status_raw is not None else ""
            if status_txt.lower() in _VALORES_VACIOS_STATUS:
                continue
            resultado[alias] = status_txt.capitalize()
        return resultado
    return None


def leer_status_excel(ruta_excel: str, log_fn=_log_default) -> dict:
    """
    Lee el Excel de Disponibilidad (una hoja por Eje/Sub-eje) y arma un
    diccionario global {ALIAS_UPPER: Status}, buscando en TODAS las hojas
    (columnas UT/UT2/Unidad/N° Numeracion + Estatus/Status/Disponibilidad). Hojas vacías,
    en texto libre, o sin esas columnas se omiten (se avisa por log_fn).

    Las hojas completamente VACÍAS (nuevo 16/09/2026) además se registran
    en status_db.ejes_sin_reportar, para que actualizar_desconexion()
    pueda escribir "Sin Reportar" en Status cuando una unidad de ese
    Eje/SubEje no aparece en ningún lado de la Disponibilidad.
    """
    try:
        xl = pd.ExcelFile(ruta_excel, engine="calamine")
    except Exception as e:
        raise ValueError(f"No se pudo leer el Excel de Disponibilidad: {e}")

    status_db = StatusDB()
    for hoja in xl.sheet_names:
        df_raw = xl.parse(hoja, header=None)
        if df_raw.empty:
            log_fn(f"  ⚠ Hoja '{hoja}' vacía, se omite.")
            par = _mapear_hoja_a_eje_subeje(hoja)
            if par:
                status_db.ejes_sin_reportar.add(par)
                log_fn(f"    → Eje/SubEje '{par[0]}/{par[1]}' sin ninguna información de Disponibilidad (Status = 'Sin Reportar').")
            continue

        datos_hoja = _parsear_hoja_status(df_raw)
        if not datos_hoja:
            log_fn(f"  ⚠ Hoja '{hoja}': no se reconoció columna de Estatus/Status/Disponibilidad, se omite.")
            continue

        for alias, status in datos_hoja.items():
            status_db[alias] = status
        log_fn(f"  ✓ Hoja '{hoja}': {len(datos_hoja)} unidad(es) con Status.")

    log_fn(f"  Total unidades con Status cargadas: {len(status_db)}")
    return status_db


def get_status(status_db: dict, alias: str):
    """Retorna el Status de la unidad si está en status_db, o None si no
    se adjuntó Excel o la unidad no aparece (en cuyo caso se conserva el
    Status que ya tenía la fila, sin borrarlo)."""
    if not status_db:
        return None
    return status_db.get(str(alias).strip().upper())


def get_status_o_sin_reportar(status_db, alias: str, eje: str, subeje: str):
    """
    Igual que get_status(), pero si la unidad no aparece en status_db Y su
    (Eje, SubEje) corresponde a una hoja de Disponibilidad completamente
    vacía (status_db.ejes_sin_reportar), retorna "Sin Reportar" en vez de
    None — nuevo 16/09/2026. Si no se adjuntó Excel de Disponibilidad
    (status_db vacío/None), retorna None como antes (no se toca la fila).
    """
    if not status_db:
        return None
    encontrado = status_db.get(str(alias).strip().upper())
    if encontrado is not None:
        return encontrado
    if (eje, subeje) in getattr(status_db, "ejes_sin_reportar", ()):
        return "Sin Reportar"
    return None


def _quitar_tildes(texto):
    """Normaliza texto: sin tildes, minúsculas, sin puntos extra."""
    texto = texto.strip().lower()
    texto = texto.replace('.', '').replace('_', ' ')
    texto = unicodedata.normalize('NFKD', texto)
    texto = ''.join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto


def normalizar_alias(alias_raw):
    """
    Convierte el alias tal como aparece en el PDF ('UT: 0089', 'R-002', 'R-032*')
    al mismo formato usado en la columna 'Unidad' del Excel ('0089', 'R-002').
    """
    alias_raw = alias_raw.strip().rstrip('*').strip()
    numeros = re.search(r'\d+', alias_raw)
    if not numeros:
        return alias_raw
    digitos = numeros.group()
    if alias_raw.upper().startswith('R'):
        return f"R-{digitos}"
    return digitos


def mapear_eje_subeje(sub_raw):
    """
    Dado el texto de sub-eje tal como aparece en el PDF (p.ej. 'TUY II', 'Acevedo'),
    retorna (eje, subeje) en el formato oficial del Excel, o (None, None) si no
    se reconoce (la unidad se agrega igual, dejando esas columnas vacías).
    """
    if not sub_raw:
        return None, None
    clave = _quitar_tildes(sub_raw)
    return MAPA_SUBEJE.get(clave, (None, None))


def leer_pdf_desconexion(ruta_pdf, log_fn=_log_default):
    """
    Lee el PDF 'Resumen de Unidades que No Han Reportado' del GTRMax.

    Retorna:
        list[dict] con llaves: unidad (alias normalizado), vin, eje, subeje,
        fecha_reporte (datetime), subeje_raw (texto original del PDF).
    """
    registros = []
    sin_parsear = []

    with pdfplumber.open(ruta_pdf) as pdf:
        for pagina in pdf.pages:
            for tabla in pagina.extract_tables():
                for fila in tabla:
                    if not fila or not fila[0] or fila[0].strip() in ('N°', ''):
                        continue
                    if len(fila) < 3:
                        continue

                    texto_unidad = (fila[1] or '').strip()
                    texto_fecha = (fila[2] or '').strip()

                    m_unidad = PATRON_UNIDAD_PDF.match(texto_unidad)
                    m_fecha = PATRON_FECHA_PDF.search(texto_fecha)

                    if not m_unidad or not m_fecha:
                        sin_parsear.append(texto_unidad or texto_fecha)
                        continue

                    alias = normalizar_alias(m_unidad.group('alias'))
                    vin = m_unidad.group('vin')
                    sub_raw = m_unidad.group('sub')
                    eje, subeje = mapear_eje_subeje(sub_raw)

                    fecha_dt = datetime.strptime(
                        f"{m_fecha.group('fecha')} {m_fecha.group('hora').upper()}",
                        "%d/%m/%Y %I:%M:%S %p"
                    )

                    registros.append({
                        'unidad': alias,
                        'vin': vin,
                        'subeje_raw': sub_raw,
                        'eje': eje,
                        'subeje': subeje,
                        'fecha_reporte': fecha_dt,
                    })

    if sin_parsear:
        log_fn(f"⚠️  {len(sin_parsear)} fila(s) del PDF no se pudieron interpretar y fueron omitidas.")

    if UNIDADES_A_OBVIAR_DEL_PDF:
        antes = len(registros)
        registros = [r for r in registros if r['unidad'] not in UNIDADES_A_OBVIAR_DEL_PDF]
        omitidas = antes - len(registros)
        if omitidas:
            log_fn(f"✔ {omitidas} unidad(es) del PDF omitidas por estar en UNIDADES_A_OBVIAR_DEL_PDF.")

    log_fn(f"✔ PDF leído: {len(registros)} unidad(es) sin reportar encontradas.")
    return registros


def leer_excel_desconexion(ruta_excel, log_fn=_log_default):
    """Abre el Excel de Desconexión y retorna (workbook, worksheet)."""
    wb = openpyxl.load_workbook(ruta_excel)
    ws = wb[wb.sheetnames[0]]
    log_fn(f"✔ Excel leído: hoja '{ws.title}' con {ws.max_row - 1} fila(s) de datos.")
    return wb, ws


def _sumar_meses(fecha, n):
    """Suma n meses de CALENDARIO a una fecha (respeta meses de 28/29/30/31 días,
    igual que lo haría Excel: si el mes destino no tiene ese día, usa su último día)."""
    mes_total = fecha.month - 1 + n
    anio = fecha.year + mes_total // 12
    mes = mes_total % 12 + 1
    dia = min(fecha.day, calendar.monthrange(anio, mes)[1])
    return fecha.replace(year=anio, month=mes, day=dia)


def _meses_calendario_completos(inicio, fin):
    """
    Cantidad de meses de CALENDARIO completos entre 'inicio' y 'fin' (fin > inicio).
    Un mes solo cuenta como completo cuando se vuelve a alcanzar exactamente el
    mismo día (y hora) del mes siguiente — igual que contarías los meses a mano
    en un calendario. Esto evita el error de usar bloques fijos de 31 días, que
    se desalinea con meses reales de 28, 29 o 30 días.
    """
    if fin <= inicio:
        return 0
    n = 0
    while _sumar_meses(inicio, n + 1) <= fin:
        n += 1
    return n


def _desglose_tiempo(fecha_actual, ultima_transmision):
    """
    Calcula el desglose EXACTO del tiempo desconectado, sin redondear ni
    adelantar nada:
      - dias:  cantidad de períodos de 24 horas EXACTAS ya cumplidos
               (si no se han cumplido 24h completas, no cuenta ni 1 día).
               Se usa mientras la desconexión sea de 31 días o menos.
      - meses: cantidad de MESES DE CALENDARIO exactos ya cumplidos (no
               bloques fijos de 31 días), para que coincida con el mes
               calendario real (28/29/30/31 días según corresponda).
      - anios: cantidad de períodos de 12 meses EXACTOS ya cumplidos.

    Nunca se redondea: el tiempo solo avanza cuando el período
    correspondiente se cumplió por completo.
    """
    total_horas = (fecha_actual - ultima_transmision).total_seconds() / 3600.0
    dias = int(total_horas // 24)       # solo días de 24h ya completados

    if dias <= 31:
        return dias, 0, 0

    meses = _meses_calendario_completos(ultima_transmision, fecha_actual)
    anios = meses // 12                 # solo años de 12 meses ya completados
    return dias, meses, anios


def calcular_tiempo_desconexion(fecha_actual, ultima_transmision):
    """
    Calcula el texto de 'Tiempo de Desconexion' según la regla de negocio,
    SIN redondear y SIN adelantar el cálculo:
      - 0 a 31 días  -> 'N Día' / 'N Días' (solo cuenta 24h ya cumplidas).
      - > 31 días    -> se convierte a meses de CALENDARIO (N Meses), y solo
                        sube de mes cuando se cumple exactamente ese mes
                        (respeta meses de 28/29/30/31 días; sin detenerse en
                        6: continúa 7, 8, 9... meses).
      - >= 12 meses  -> se convierte a años (N Año / N Años), solo cuando
                        se cumplen 12 meses completos.
    """
    dias, meses, anios = _desglose_tiempo(fecha_actual, ultima_transmision)

    if dias <= 31:
        return f"{dias} Día" if dias == 1 else f"{dias} Días"

    if meses < 12:
        return f"{meses} Mes" if meses == 1 else f"{meses} Meses"

    return f"{anios} Año" if anios == 1 else f"{anios} Años"


def _parsear_fecha_actual(texto):
    """
    Convierte el texto de 'fecha actual' a un datetime correcto para los
    cálculos exactos de horas/días/meses.

    IMPORTANTE (corrige el error de "falta 1 día"):
    Si el texto trae solo la fecha (DD/MM/AAAA), NO se asume medianoche
    (00:00:00), porque eso resta horas reales y hace que una unidad
    aparezca con 1 día menos de desconexión de los que realmente tiene
    (por ejemplo: última transmisión 28/07 09:50 AM y fecha actual 02/08
    debe dar 5 días completos, no 4).

    En su lugar se usa el FINAL de ese día (23:59:59), es decir, se asume
    que ese día ya transcurrió por completo. Así el resultado es siempre
    igual a la diferencia de fechas de calendario, sin importar a qué hora
    del día se ejecute el script.

    Si el texto ya trae una hora específica (DD/MM/AAAA HH:MM o
    DD/MM/AAAA HH:MM:SS), se usa esa hora tal cual, sin modificarla.
    """
    texto = texto.strip()

    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M"):
        try:
            return datetime.strptime(texto, fmt)
        except ValueError:
            continue

    fecha = datetime.strptime(texto, "%d/%m/%Y")
    return fecha.replace(hour=23, minute=59, second=59, microsecond=0)


def unidad_tiene_24h_cumplidas(fecha_actual, ultima_transmision):
    """
    Retorna True solo si la unidad ya cumplió al menos 24 horas EXACTAS de
    desconexión. Las unidades con menos de 24h se deben ignorar/omitir del
    reporte (aún no se consideran 'desconectadas').
    """
    dias, _, _ = _desglose_tiempo(fecha_actual, ultima_transmision)
    return dias >= 1


def _capturar_estilos_fila(ws, fila, num_columnas):
    """Devuelve una lista de estilos (font/fill/border/alignment/number_format)
    copiados de una fila, para poder reaplicarlos luego de reordenar los datos."""
    estilos = []
    for col in range(1, num_columnas + 1):
        celda = ws.cell(row=fila, column=col)
        estilos.append({
            'font': copy.copy(celda.font),
            'fill': copy.copy(celda.fill),
            'border': copy.copy(celda.border),
            'alignment': copy.copy(celda.alignment),
            'number_format': celda.number_format,
        })
    return estilos


def _aplicar_estilos_fila(ws, fila, estilos):
    """Aplica a una fila la lista de estilos capturada con _capturar_estilos_fila."""
    for col, estilo in enumerate(estilos, start=1):
        celda = ws.cell(row=fila, column=col)
        celda.font = estilo['font']
        celda.fill = estilo['fill']
        celda.border = estilo['border']
        celda.alignment = estilo['alignment']
        celda.number_format = estilo['number_format']


def _clave_orden_alfabetico(fila):
    """
    Orden alfabético: Eje -> Sub Eje -> Unidad (insensible a mayúsculas/tildes).
    Las filas sin Eje/Sub Eje reconocido se muestran al final de su grupo.
    """
    eje = _quitar_tildes(fila['eje'] or '')
    subeje = _quitar_tildes(fila['subeje'] or '')
    unidad = _quitar_tildes(str(fila['unidad'] or ''))
    return (eje == '', eje, subeje == '', subeje, unidad)


def crear_excel_desconexion_nuevo(responsable="", corte="", log_fn=_log_default):
    """
    Crea un Excel de Desconexión NUEVO (sin plantilla previa), con la misma
    estructura de columnas que el formato original. Se usa cuando el usuario
    no tiene un archivo existente para actualizar.

    Retorna (Workbook, Worksheet) listos para que actualizar_desconexion()
    los llene con las unidades del PDF.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Desconexion"

    encabezados = {
        COL_RESPONSABLE: "Responsable", COL_FECHA: "Fecha", COL_CORTE: "Corte",
        COL_ULTIMA_TX: "Ultima Transmision", COL_EJE: "Eje", COL_SUBEJE: "Sub Eje",
        COL_UNIDAD: "Unidad", COL_TIEMPO_DESC: "Tiempo de Desconexion",
        COL_COND_GPS: "Condicion de GPS", COL_STATUS: "Status",
    }

    fill_header = PatternFill(start_color="1B3A6B", end_color="1B3A6B", fill_type="solid")
    font_header = Font(color="FFFFFF", bold=True, size=11)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="B0B0B0")
    border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, texto in encabezados.items():
        celda = ws.cell(row=1, column=col, value=texto)
        celda.font = font_header
        celda.fill = fill_header
        celda.alignment = align_center
        celda.border = border_thin
        ws.column_dimensions[celda.column_letter].width = 20

    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"

    log_fn("✔ Creado un Excel de Desconexión nuevo (sin plantilla previa).")
    return wb, ws


def _estilo_banda_nueva(color_hex):
    """
    Genera los estilos de una fila de datos (fondo, bordes, formato de fecha)
    para cuando se está creando un Excel nuevo y no hay ninguna fila existente
    de la que copiar el formato.
    """
    fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    align = Alignment(horizontal="center", vertical="center")
    formatos_fecha = {COL_FECHA: "dd/mm/yyyy", COL_ULTIMA_TX: "dd/mm/yyyy hh:mm:ss"}

    estilos = []
    for col in range(1, 11):
        estilos.append({
            'font': Font(size=10),
            'fill': fill,
            'border': border,
            'alignment': align,
            'number_format': formatos_fecha.get(col, "General"),
        })
    return estilos


def actualizar_desconexion(ruta_pdf, ruta_excel, fecha_actual, carpeta_salida, log_fn=_log_default,
                            responsable_nuevo="", corte_nuevo="", status_db=None):
    """
    Función principal: combina PDF + Excel y genera el Excel actualizado.

    Reglas aplicadas:
      1. Actualiza 'Ultima Transmision' / 'Tiempo de Desconexion' con los datos del PDF.
      2. Agrega como fila nueva cualquier unidad del PDF que no exista en el Excel.
      3. Si una unidad YA estaba en el Excel pero ya NO aparece en el PDF actual,
         se elimina del resultado (significa que volvió a transmitir / ya no
         está desconectada) — salvo que esté en UNIDADES_A_OBVIAR_DEL_PDF.
      4. Omite (no incluye en el resultado) toda unidad con MENOS de 24 horas
         exactas de desconexión (aún no se considera desconectada).
      5. Incluye TODAS las demás unidades sin importar cuántos meses/años lleven.
      6. Ordena el resultado alfabéticamente por Eje -> Sub Eje -> Unidad.
      7. Si se adjunta un Excel de Disponibilidad (status_db), actualiza la
         columna 'Status' de cada unidad con el valor encontrado ahí; si la
         unidad no aparece en ese Excel, conserva el Status que ya tenía.

    Parámetros:
        ruta_pdf (str): ruta al PDF de GTRMax "Resumen de Unidades...".
        ruta_excel (str | None): ruta al Excel de Desconexión a actualizar.
            Si es None, se crea un Excel nuevo desde cero (usa
            responsable_nuevo / corte_nuevo para esas columnas).
        fecha_actual (datetime): fecha (y hora) que se usa como "hoy" para
            recalcular los tiempos de desconexión.
        carpeta_salida (str): carpeta donde se guarda el archivo resultante.
        log_fn (callable): función de log (por defecto, print).
        responsable_nuevo (str): valor de la columna Responsable, solo se usa
            cuando ruta_excel es None (Excel nuevo).
        corte_nuevo (str): valor de la columna Corte, solo se usa cuando
            ruta_excel es None (Excel nuevo).
        status_db (dict | None): {ALIAS: Status} devuelto por
            leer_status_excel(), o None si no se adjuntó Excel de
            Disponibilidad (en ese caso Status se comporta igual que antes).

    Retorna:
        str: ruta del archivo Excel generado.
    """
    registros_pdf = leer_pdf_desconexion(ruta_pdf, log_fn)

    crear_nuevo = not ruta_excel
    if crear_nuevo:
        wb, ws = crear_excel_desconexion_nuevo(responsable_nuevo, corte_nuevo, log_fn)
        num_columnas = 10
        responsable_global = responsable_nuevo
        corte_global = corte_nuevo
        estilo_fila_par = _estilo_banda_nueva("FFFFFF")
        estilo_fila_impar = _estilo_banda_nueva("F2F2F2")
        filas_datos = []
        fila_max_original = ws.max_row
    else:
        wb, ws = leer_excel_desconexion(ruta_excel, log_fn)

        # Renombrar el encabezado de esta columna aunque el Excel existente
        # traiga el nombre viejo ("Cond GPS").
        ws.cell(row=1, column=COL_COND_GPS).value = "Condicion de GPS"

        num_columnas = ws.max_column
        fila_max_original = ws.max_row

        # Valores "globales" que se repiten en todas las filas del Excel
        # (Responsable y Corte se mantienen igual a como estaban).
        responsable_global = ws.cell(row=FILA_INICIAL_DATOS, column=COL_RESPONSABLE).value
        corte_global = ws.cell(row=FILA_INICIAL_DATOS, column=COL_CORTE).value

        # Plantillas de estilo (banda par/impar) capturadas ANTES de tocar la
        # hoja, para poder reaplicarlas tal cual luego de reordenar los datos.
        estilo_fila_par = _capturar_estilos_fila(ws, FILA_INICIAL_DATOS, num_columnas)
        estilo_fila_impar = _capturar_estilos_fila(ws, FILA_INICIAL_DATOS + 1, num_columnas)

        # 1) Cargar a memoria todas las filas de datos existentes -----------
        filas_datos = []
        for fila in range(FILA_INICIAL_DATOS, fila_max_original + 1):
            unidad = ws.cell(row=fila, column=COL_UNIDAD).value
            if unidad is None or str(unidad).strip() == '':
                continue
            filas_datos.append({
                'unidad': str(unidad).strip(),
                'eje': ws.cell(row=fila, column=COL_EJE).value,
                'subeje': ws.cell(row=fila, column=COL_SUBEJE).value,
                'ultima_tx': ws.cell(row=fila, column=COL_ULTIMA_TX).value,
                'cond_gps': ws.cell(row=fila, column=COL_COND_GPS).value,
                'status': ws.cell(row=fila, column=COL_STATUS).value,
            })

    # 1.5) Quitar del Excel existente las unidades que YA NO aparecen en el
    #      PDF actual (se reconectaron) — salvo las que estén en la lista
    #      manual de unidades a obviar, esas se mantienen intactas siempre.
    unidades_pdf = {reg['unidad'] for reg in registros_pdf}
    reconectadas = 0
    if filas_datos:
        filas_filtradas = []
        for f in filas_datos:
            if f['unidad'] in unidades_pdf or f['unidad'] in UNIDADES_A_OBVIAR_DEL_PDF:
                filas_filtradas.append(f)
            else:
                reconectadas += 1
        filas_datos = filas_filtradas

    # 2) Aplicar los datos del PDF: actualizar existentes / agregar nuevas ---
    actualizadas = 0
    agregadas = 0

    for reg in registros_pdf:
        unidad = reg['unidad']
        fecha_reporte = reg['fecha_reporte']
        coincidencias = [f for f in filas_datos if f['unidad'] == unidad]

        if coincidencias:
            for f in coincidencias:
                if f['ultima_tx'] == fecha_reporte:
                    # Misma fecha que ya tenía -> sigue desconectada; se
                    # mantiene y el tiempo se recalcula con la fecha actual
                    # (esto suma automáticamente 1 día más).
                    pass
                else:
                    # El PDF trae una fecha de última transmisión distinta ->
                    # se actualiza con el dato más reciente del PDF.
                    f['ultima_tx'] = fecha_reporte
                actualizadas += 1
        else:
            filas_datos.append({
                'unidad': unidad,
                'eje': reg['eje'],       # None si no se reconoce -> queda vacío
                'subeje': reg['subeje'],  # None si no se reconoce -> queda vacío
                'ultima_tx': fecha_reporte,
                'cond_gps': None,
                'status': None,
            })
            agregadas += 1

    # 3) Recalcular tiempo de desconexión para TODAS las filas y omitir
    #    las que tengan MENOS de 24 horas exactas de desconexión ------------
    filas_finales = []
    excluidas = 0
    for f in filas_datos:
        if f['ultima_tx'] is None:
            continue
        if not unidad_tiene_24h_cumplidas(fecha_actual, f['ultima_tx']):
            excluidas += 1
            continue
        f['tiempo_desc'] = calcular_tiempo_desconexion(fecha_actual, f['ultima_tx'])
        filas_finales.append(f)

    # 4) Ordenar alfabéticamente: Eje -> Sub Eje -> Unidad -------------------
    filas_finales.sort(key=_clave_orden_alfabetico)

    # 5) Reescribir la hoja desde cero con el orden y los datos finales -----
    if fila_max_original >= FILA_INICIAL_DATOS:
        ws.delete_rows(FILA_INICIAL_DATOS, fila_max_original - FILA_INICIAL_DATOS + 1)

    fecha_col_b = fecha_actual.replace(hour=0, minute=0, second=0, microsecond=0)
    status_actualizados = 0

    for i, f in enumerate(filas_finales):
        fila = FILA_INICIAL_DATOS + i
        estilo = estilo_fila_par if i % 2 == 0 else estilo_fila_impar
        _aplicar_estilos_fila(ws, fila, estilo)

        status_nuevo = get_status_o_sin_reportar(status_db, f['unidad'], f['eje'], f['subeje'])
        if status_nuevo is not None:
            f['status'] = status_nuevo
            status_actualizados += 1

        f['cond_gps'] = _condicion_gps(f['unidad'])

        ws.cell(row=fila, column=COL_RESPONSABLE).value = responsable_global
        ws.cell(row=fila, column=COL_FECHA).value = fecha_col_b
        ws.cell(row=fila, column=COL_CORTE).value = corte_global
        ws.cell(row=fila, column=COL_ULTIMA_TX).value = f['ultima_tx']
        ws.cell(row=fila, column=COL_EJE).value = f['eje']
        ws.cell(row=fila, column=COL_SUBEJE).value = f['subeje']
        ws.cell(row=fila, column=COL_UNIDAD).value = f['unidad']
        ws.cell(row=fila, column=COL_TIEMPO_DESC).value = f['tiempo_desc']
        ws.cell(row=fila, column=COL_COND_GPS).value = f['cond_gps']
        ws.cell(row=fila, column=COL_STATUS).value = f['status']

    log_fn(f"✔ Unidades actualizadas: {actualizadas}")
    log_fn(f"✔ Unidades nuevas agregadas: {agregadas}")
    log_fn(f"✔ Unidades removidas por ya no aparecer en el PDF (reconectadas): {reconectadas}")
    log_fn(f"✔ Unidades omitidas por tener menos de 24h de desconexión: {excluidas}")
    if status_db:
        log_fn(f"✔ Unidades con Status actualizado desde el Excel de Disponibilidad: {status_actualizados}")
    log_fn(f"✔ Total de filas en el archivo final: {len(filas_finales)} (ordenadas alfabéticamente)")

    return guardar_excel(wb, carpeta_salida, log_fn)


def guardar_excel(wb, carpeta_salida, log_fn=_log_default):
    """Guarda el workbook actualizado con nombre Data_de_desconexion_DDMMYYYY.xlsx"""
    os.makedirs(carpeta_salida, exist_ok=True)
    nombre = f"Data_de_desconexion_{datetime.now().strftime('%d%m%Y')}.xlsx"
    ruta_salida = os.path.join(carpeta_salida, nombre)
    wb.save(ruta_salida)
    log_fn(f"✔ Archivo guardado en: {ruta_salida}")
    return ruta_salida


# ─────────────────────────────────────────────────────────────────────────────
#  Punto de entrada (modo consola / script independiente)
# ─────────────────────────────────────────────────────────────────────────────
def _pausar_antes_de_cerrar():
    """Evita que la ventana de consola se cierre sola (doble clic en Windows)."""
    try:
        input("\nPresiona ENTER para cerrar esta ventana...")
    except (EOFError, KeyboardInterrupt):
        pass


def _pedir_ruta(mensaje, debe_existir=True):
    """Pide una ruta por consola, quita comillas que Windows agrega al arrastrar
    archivos, y valida que exista."""
    while True:
        ruta = input(mensaje).strip().strip('"').strip("'")
        if not debe_existir or os.path.isfile(ruta):
            return ruta
        print(f"⚠ No se encontró el archivo: {ruta}\n  Intenta de nuevo (o arrastra el archivo a esta ventana).")


def _pedir_fecha_actual():
    """Pide la fecha actual (o usa FECHA_ACTUAL / hoy si el usuario no escribe nada)."""
    if FECHA_ACTUAL:
        return _parsear_fecha_actual(FECHA_ACTUAL)

    texto = input(
        "Fecha actual (DD/MM/AAAA) — presiona ENTER para usar la fecha de hoy: "
    ).strip()
    if not texto:
        return datetime.now()
    return _parsear_fecha_actual(texto)


def _modo_interactivo():
    """Se activa cuando el script se abre con doble clic (sin argumentos)."""
    print("=" * 70)
    print(" ACTUALIZAR DESCONEXIÓN — ControlPlus / MoussaCorp / TransMiranda")
    print("=" * 70)

    ruta_pdf = _pedir_ruta("\nRuta del PDF (Resumen de Unidades sin reportar): ")
    ruta_excel = _pedir_ruta("Ruta del Excel de Desconexión a actualizar: ")

    carpeta_por_defecto = os.path.dirname(os.path.abspath(ruta_excel))
    carpeta_salida = input(
        f"Carpeta de salida [ENTER = {carpeta_por_defecto}]: "
    ).strip().strip('"').strip("'") or carpeta_por_defecto

    fecha_actual_final = _pedir_fecha_actual()

    print()
    actualizar_desconexion(ruta_pdf, ruta_excel, fecha_actual_final, carpeta_salida)


def _modo_argumentos():
    """Modo clásico: python actualizar_desconexion.py pdf.xlsx excel.xlsx [salida] [fecha]"""
    ruta_pdf_arg = sys.argv[1]
    ruta_excel_arg = sys.argv[2]
    carpeta_salida_arg = sys.argv[3] if len(sys.argv) > 3 else os.path.dirname(os.path.abspath(ruta_excel_arg))

    fecha_texto = sys.argv[4] if len(sys.argv) > 4 else FECHA_ACTUAL
    fecha_actual_final = _parsear_fecha_actual(fecha_texto) if fecha_texto else datetime.now()

    actualizar_desconexion(ruta_pdf_arg, ruta_excel_arg, fecha_actual_final, carpeta_salida_arg)


if __name__ == '__main__':
    try:
        if len(sys.argv) >= 3:
            _modo_argumentos()
        else:
            _modo_interactivo()
    except Exception:
        # Si algo falla, mostrar el error completo en vez de que la ventana
        # se cierre antes de poder leerlo.
        import traceback
        print("\n❌ Ocurrió un error:\n")
        traceback.print_exc()
    finally:
        _pausar_antes_de_cerrar()
