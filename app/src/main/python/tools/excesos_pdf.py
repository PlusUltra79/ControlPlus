"""
Excesos de Velocidad desde PDF — GTRMax
----------------------------------------
Lee el PDF "Notificaciones de alarmas" del sistema GTRMax.
Filtra SOLO las filas de velocidad, aceptando como equivalentes las
etiquetas "Agente de Velocidad" y "Exceso de Velocidad" (el PDF puede
traer cualquiera de las dos según la versión del reporte).
Por cada unidad, conserva únicamente el registro de mayor velocidad.

El saludo del mensaje ahora depende de la hora que indica el usuario:
  04:00 - 11:59 → "Buenos Días"
  12:00 - 18:59 → "Buenas Tardes"
  19:00 - 03:59 → "Buenas Noches"

El CTS ya no se deja fijo en "S.I." — se busca la unidad en un Excel de
Disponibilidad (una hoja por Eje/Sub-eje, ver leer_cts_excel()). Si la
unidad no aparece en el Excel (o no se adjuntó Excel), se usa "S.I." como
antes.

Formato del mensaje (caso general):
  Buenas Tardes
  * Eje *Blv* Sub Eje *Acevedo*
  * Unidad *0645* - Vin *0684*
  * CTS: *S.I.*
  *Incidencia:* Exceso de velocidad
  *Incumplimiento:* *93 km/h* en *autopista*

Casos especiales en la línea "Eje" (ver formatear_linea_eje()):
  - Pz no tiene sub-eje propio: "* Eje *Pz*" (sin "Sub Eje").
  - Serv Cont se muestra combinado: "* Eje *Serv. Cont*" (sin "Sub Eje").
  - Serv Esp se muestra combinado: "* Eje *Serv Esp*" (sin "Sub Eje",
    sin punto — a diferencia de Serv Cont).

Excepción manual por unidad (UNIDADES_FORZADAS): la unidad R-036 se
clasifica SIEMPRE como Serv Esp, sin importar qué eje indique el PDF.

Requiere: pip install pdfplumber pandas python-calamine
Uso: python excesos_pdf.py
"""

import os, re
import numbers
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from datetime import datetime

# ── VIN Database ───────────────────────────────────────────────────────────────
VIN_DB = {
    "R-003":"4716","R-022":"4919","R-045":"3916","R-046":"4625",
    "R-047":"9924","R-048":"4791","R-053":"8126","R-060":"9867",
    "R-066":"2671","R-069":"8396","0219":"8792","0349":"0589",
    "0489":"8627","0491":"8644","0566":"6324","0598":"0415",
    "0613":"0280","0615":"2295","0647":"0693","0654":"0722",
    "0664":"0745","R-010":"5944","0352":"0624","0452":"4436",
    "0599":"0396","0643":"0705","0645":"0684","0656":"0723",
    "0665":"0744","0680":"0179","0721":"0101","0722":"0052",
    "R-006":"9616","R-007":"2273","R-009":"7225","R-052":"8252",
    "R-058":"2109","0426":"4361","0522":"0464","0641":"0703",
    "0657":"0712","0666":"0713","0700":"0031","R-005":"9299",
    "R-019":"5738","0416":"0132","0465":"4316","0579":"0426",
    "0616":"2244","R-049":"4622","R-050":"4874","0124":"0947",
    "0343":"0573","0481":"0781","0482":"4276","0582":"3270",
    "0609":"6312","0629":"8060","0636":"0682","0642":"0691",
    "0646":"0690","0648":"0692","R-008":"8319","R-034":"6159",
    "R-056":"0039","0477":"4434","0667":"0725","R-012":"4649",
    "R-013":"4955","R-014":"4742","R-031":"4859","0429":"4326",
    "0495":"8661","0497":"9519","0500":"9536","0512":"8436",
    "0516":"7824","0531":"0404","0546":"6384","0581":"0423",
    "0597":"0428","0685":"0257","R-025":"4954","R-033":"4665",
    "R-051":"3750","0436":"4458","0511":"9921","0612":"6362",
    "0632":"7938","R-039":"4778","R-040":"0166","R-041":"4765",
    "R-062":"1109","0140":"1182","0424":"4343","0462":"4412",
    "0463":"4366","0485":"4413","0536":"0427","0624":"7851",
    "0625":"7850","0678":"0242","0705":"0108","R-042":"7819",
    "R-043":"4769","R-044":"2905","R-054":"3606","R-055":"3949",
    "R-057":"7799","R-063":"4704","0417":"4470","0445":"4273",
    "0450":"4270","0454":"4505","0614":"2277","R-001*":"4918",
    "R-016":"4728","R-017":"2120","R-018":"1419","R-020":"8092",
    "R-021":"3679","R-023":"4118","R-024":"3419","R-026":"4923",
    "R-027":"4718","R-029":"4043","R-061":"1481","R-067":"8148",
    "0425":"4352","0448":"4370","0480":"4414","0484":"4344",
    "0503":"7815","0583":"0383","0608":"6346","0637":"0696",
    "0638":"5153","0446":"4473","0466":"4489","0592":"2186",
    "0673":"6627","0690":"0276","0697":"0275","R-011":"4650",
    "0505":"7852","0514":"9545","0533":"6307","0538":"0393",
    "0557":"0385","0562":"6353","0576":"0454","0591":"0460",
    "0594":"0452","0671":"6481","0672":"6490","0683":"0281",
    "0702":"0016","0712":"0041","R-064":"6656",
}

# Mapeo sub-eje → Eje (tal como aparece en el PDF, en mayúsculas/minúsculas varias)
SUBEJE_A_EJE = {
    # Am
    "NORTE":"Am","CENTRO":"Am","SUR":"Am",
    "Norte":"Am","Centro":"Am","Sur":"Am",
    # Blv
    "ACEVEDO":"Blv","BRION":"Blv","BRIÓN":"Blv",
    "BUROZ":"Blv","PÁEZ":"Blv","PAEZ":"Blv",
    "PEDRO GUAL":"Blv","PÁEZ ANDRÉS BELLO":"Blv",
    "Acevedo":"Blv","Brion":"Blv","Brión":"Blv",
    "Buroz":"Blv","Páez":"Blv","Paez":"Blv",
    "Pedro Gual":"Blv","Pedro gual":"Blv",
    # Met
    "MET":"Met","METRO":"Met","SUCRE":"Met",
    "Met":"Met","Metro":"Met","Sucre":"Met",
    # Ocm
    "TUY I":"Ocm","TUY II":"Ocm",
    "Tuy I":"Ocm","Tuy II":"Ocm",
    # Pz
    "PZ":"Pz","PLAZA ZAMORA":"Pz","PLAZA SUCRE":"Pz",
    "Pz":"Pz","Plaza Zamora":"Pz","Plaza Sucre":"Pz",
    # Serv
    "ESP":"Serv","CONT":"Serv",
    "Esp":"Serv","Cont":"Serv",
}

# Excepción manual por unidad: fuerza (eje, sub_eje) sin importar lo que
# indique el PDF. La unidad R-036 pertenece siempre a Serv Esp, aunque el
# PDF la reporte como si fuera de Pz (mismo criterio que UNIDADES_FORZADAS
# en ultimas_ut_movimiento.py). Si aparecen más unidades con esta misma
# situación, se agregan aquí.
UNIDADES_FORZADAS = {
    "R-036": ("Serv", "Esp"),
}


# Etiquetas equivalentes que puede traer la columna "Agente" en el PDF.
# El GTRMax a veces usa "Agente de Velocidad" y otras veces "Exceso de
# Velocidad" para referirse al mismo tipo de alarma. Ambas se tratan igual.
AGENTES_VELOCIDAD_VALIDOS = {
    "agente de velocidad",
    "exceso de velocidad",
}


def es_agente_velocidad(valor_agente: str) -> bool:
    """
    True si el valor de la columna Agente corresponde a una alarma de
    velocidad, sin importar si el PDF usa "Agente de Velocidad" o
    "Exceso de Velocidad" (comparación insensible a mayúsculas y espacios).
    """
    normalizado = str(valor_agente).strip().lower()
    if normalizado in AGENTES_VELOCIDAD_VALIDOS:
        return True
    # Respaldo: cualquier variante que contenga la palabra "velocidad"
    # (cubre pequeñas diferencias de redacción no listadas arriba)
    return "velocidad" in normalizado


def get_vin(alias: str) -> str:
    key = alias.strip().upper()
    for k, v in VIN_DB.items():
        if k.upper() == key:
            return v
    return "S/N"


def _formatear_alias_numerico(valor) -> str:
    """
    Corrige la misma fragilidad detectada en Actualizar Desconexión
    (16/09/2026, unidad 0140 de Ocm): si la columna de alias de una hoja de
    Disponibilidad es TOTALMENTE numérica, python-calamine la lee como
    número y pierde los ceros a la izquierda ("0140" -> 140.0). Se
    reconstruye como texto de 4 dígitos (formato fijo de alias numérico
    usado en todo el proyecto) cuando el valor llega como int/float.
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


def limpiar_alias(raw) -> str:
    return re.sub(r"(?i)^ut:\s*", "", _formatear_alias_numerico(raw)).strip()


# ── Hora manual → saludo ─────────────────────────────────────────────────────

def parsear_hora_24h(texto: str):
    """
    Convierte la hora que escribe el usuario a (hora, minuto) en 24h, para
    poder calcular el saludo. Admite 24h ("21:15") o 12h con am/pm
    ("08:02 pm", "8:02 PM", "08:02 p.m.").
    """
    limpio = texto.strip()

    m = re.match(r'^(\d{1,2}):(\d{2})\s*([ap])\.?\s*m\.?$', limpio, re.IGNORECASE)
    if m:
        hora = int(m.group(1)) % 12
        minuto = int(m.group(2))
        if not (0 <= hora <= 11 and 0 <= minuto <= 59):
            raise ValueError(f"Formato de hora inválido: '{texto}'.")
        if m.group(3).lower() == "p":
            hora += 12
        return hora, minuto

    m = re.match(r'^(\d{1,2}):(\d{2})$', limpio)
    if m:
        hora = int(m.group(1))
        minuto = int(m.group(2))
        if not (0 <= hora <= 23 and 0 <= minuto <= 59):
            raise ValueError(f"Formato de hora inválido: '{texto}'.")
        return hora, minuto

    raise ValueError(
        f"Formato de hora inválido: '{texto}'.\n"
        f"Usa por ejemplo: 21:15  o  08:02 pm"
    )


def saludo_por_hora(hora: int, minuto: int) -> str:
    """
    04:00–11:59 → Buenos Días · 12:00–18:59 → Buenas Tardes ·
    19:00–03:59 → Buenas Noches (cubre también la madrugada, ya que el
    rango pedido "19:00-00:00" no tiene sentido operativo dejar sin saludo).
    """
    minutos_totales = hora * 60 + minuto
    if 4 * 60 <= minutos_totales <= 11 * 60 + 59:
        return "Buenos Días"
    if 12 * 60 <= minutos_totales <= 18 * 60 + 59:
        return "Buenas Tardes"
    return "Buenas Noches"


# ── CTS desde Excel de Disponibilidad ────────────────────────────────────────

# Valores de la columna CTS que equivalen a "sin CTS asignado" y por lo
# tanto deben mostrarse como "S.I." en el mensaje (igual que sin Excel).
_CTS_VACIOS = {"", "nan", "sin cts", "s.i", "s.i.", "s/n", "-", "none"}


def _celda_alias_valida(texto: str) -> bool:
    """True si el encabezado de columna corresponde a la unidad (UT/UT2/
    Unidad/N° Numeracion, sin importar mayúsculas ni tildes)."""
    t = texto.strip().lower()
    return t in {"ut", "ut2", "unidad"} or "numeracion" in t


def _celda_cts_valida(texto: str) -> bool:
    """True si el encabezado de columna corresponde al CTS."""
    return "cts" in texto.strip().lower()


def _detectar_columnas_tabla(fila_encabezado: list):
    idx_alias, idx_cts = None, None
    for i, celda in enumerate(fila_encabezado):
        texto = str(celda) if celda is not None else ""
        if idx_alias is None and _celda_alias_valida(texto):
            idx_alias = i
        if idx_cts is None and _celda_cts_valida(texto):
            idx_cts = i
    return idx_alias, idx_cts


def _parsear_hoja_tabular(df_raw: "pd.DataFrame"):
    """
    Intenta leer la hoja como tabla (columnas UT/UT2/Unidad/N° Numeracion +
    CTS/Cts). Busca la fila de encabezado entre las primeras 3 filas, ya
    que no todas las hojas del Excel usan exactamente el mismo layout.
    Retorna None si no se reconoce ninguna fila de encabezado válida.
    """
    for fila_idx in range(min(3, len(df_raw))):
        fila = df_raw.iloc[fila_idx].tolist()
        idx_alias, idx_cts = _detectar_columnas_tabla(fila)
        if idx_alias is None or idx_cts is None:
            continue

        resultado = {}
        for r in range(fila_idx + 1, len(df_raw)):
            fila_datos = df_raw.iloc[r].tolist()
            if idx_alias >= len(fila_datos) or idx_cts >= len(fila_datos):
                continue
            alias_raw = fila_datos[idx_alias]
            if alias_raw is None or not str(alias_raw).strip():
                continue
            alias = limpiar_alias(alias_raw).upper()
            cts_raw = fila_datos[idx_cts]
            cts_txt = str(cts_raw).strip() if cts_raw is not None else ""
            if cts_txt.lower() in _CTS_VACIOS:
                continue
            resultado[alias] = cts_txt
        return resultado
    return None


def _parsear_hoja_libre(df_raw: "pd.DataFrame"):
    """
    Hojas en formato libre (sin tabla), ej. bloques de texto tipo:
        Ut:0718
        Cts: Edison Gómez
    en vez de columnas. Recorre la primera columna buscando pares
    Ut:/Cts: consecutivos.
    """
    if df_raw.shape[1] == 0:
        return {}
    lineas = [
        str(v).strip() if v is not None and str(v).strip().lower() != "nan" else ""
        for v in df_raw.iloc[:, 0].tolist()
    ]
    resultado = {}
    alias_pendiente = None
    for linea in lineas:
        m_ut = re.match(r"(?i)^ut:?\s*(\S+)", linea)
        if m_ut:
            alias_pendiente = m_ut.group(1).strip().upper()
            continue
        m_cts = re.match(r"(?i)^cts:?\s*(.*)$", linea)
        if m_cts and alias_pendiente:
            nombre = m_cts.group(1).strip()
            if nombre and nombre.lower() not in _CTS_VACIOS:
                resultado[alias_pendiente] = nombre
            alias_pendiente = None
    return resultado


def leer_cts_excel(ruta_excel: str, log_fn=print) -> dict:
    """
    Lee el Excel de Disponibilidad (una hoja por Eje/Sub-eje, ej. 'Acevedo',
    'Brion', 'Met', 'Serv Esp'...) y arma un diccionario global
    {ALIAS_UPPER: nombre_CTS}, buscando en TODAS las hojas sin importar en
    cuál esté cada unidad (el nombre de la hoja es solo referencial).

    Soporta dos formatos de hoja, detectados automáticamente:
      - Tabular: columnas UT / UT2 / Unidad / N° Numeracion + CTS / Cts
        (la fila de encabezado puede estar en cualquiera de las primeras
        3 filas, y las columnas pueden venir en cualquier orden).
      - Libre: bloques de texto "Ut:XXXX" seguido de "Cts: Nombre"
        (ej. hoja 'Buroz' del archivo real de Disponibilidad).

    Hojas vacías o sin ninguno de los dos formatos reconocibles se
    omiten (se avisa por log_fn, no se interrumpe el proceso).
    """
    try:
        xl = pd.ExcelFile(ruta_excel, engine="calamine")
    except Exception as e:
        raise ValueError(f"No se pudo leer el Excel de CTS: {e}")

    cts_db = {}
    for hoja in xl.sheet_names:
        df_raw = xl.parse(hoja, header=None)
        if df_raw.empty:
            log_fn(f"  ⚠ Hoja '{hoja}' vacía, se omite.")
            continue

        datos_hoja = _parsear_hoja_tabular(df_raw)
        if not datos_hoja:
            datos_hoja = _parsear_hoja_libre(df_raw)

        if not datos_hoja:
            log_fn(f"  ⚠ Hoja '{hoja}': no se reconoció ningún formato de CTS, se omite.")
            continue

        for alias, cts in datos_hoja.items():
            if alias in cts_db and cts_db[alias] != cts:
                log_fn(
                    f"  ⚠ Unidad {alias}: CTS distinto en más de una hoja "
                    f"('{cts_db[alias]}' → '{cts}' en '{hoja}'), se usa el último."
                )
            cts_db[alias] = cts
        log_fn(f"  ✓ Hoja '{hoja}': {len(datos_hoja)} unidad(es) con CTS.")

    log_fn(f"  Total unidades con CTS cargadas: {len(cts_db)}")
    return cts_db


def formatear_cts(nombre: str) -> str:
    """
    Formatea el nombre del CTS como 'Nombre I.' (nombre completo + inicial
    del apellido con punto), ej. 'Johan Sosa' -> 'Johan S.',
    'PABLO ROSALES' -> 'Pablo R.'. Si el nombre trae una sola palabra, se
    deja tal cual (solo con mayúscula inicial).
    """
    partes = nombre.strip().split()
    if not partes:
        return nombre.strip()
    primero = partes[0].capitalize()
    if len(partes) == 1:
        return primero
    inicial = partes[1][0].upper()
    return f"{primero} {inicial}."


def get_cts(cts_db: dict, alias: str) -> str:
    """Busca el CTS de una unidad en el diccionario cargado por
    leer_cts_excel(). Si no hay Excel, o la unidad no aparece, o su CTS
    está vacío, retorna 'S.I.' (mismo comportamiento que antes). Si se
    encuentra, se muestra como 'Nombre I.' (ver formatear_cts)."""
    if not cts_db:
        return "S.I."
    valor = cts_db.get(limpiar_alias(alias).upper())
    return formatear_cts(valor) if valor else "S.I."


def extraer_subeje_eje(descripcion: str):
    """
    Extrae sub-eje desde la descripción.
    Patrón: "(VIN: 0684 ACEVEDO)" → sub_eje="Acevedo", eje="Blv"
    """
    # Buscar todo lo que esté después del VIN dentro del paréntesis
    m = re.search(r"\(VIN:\s*\d+\s+([^)]+)\)", descripcion, re.IGNORECASE)
    if not m:
        m = re.search(r"\(Vin:\s*\d+\s+([^)]+)\)", descripcion, re.IGNORECASE)
    if not m:
        return "", ""

    raw_sub = m.group(1).strip()

    # Intentar buscar en el mapa primero con el valor original
    eje = SUBEJE_A_EJE.get(raw_sub, "")
    if not eje:
        # Buscar en uppercase
        eje = SUBEJE_A_EJE.get(raw_sub.upper(), "")
    if not eje:
        # Buscar con title case
        eje = SUBEJE_A_EJE.get(raw_sub.title(), "")

    # Formatear sub-eje con title case para el mensaje
    sub_eje_display = raw_sub.title()

    return sub_eje_display, eje


def formatear_linea_eje(eje: str, sub_eje: str) -> str:
    """
    Arma la línea 'Eje' del mensaje. Tres casos especiales sin la palabra
    'Sub Eje':
      - Pz: no tiene sub-eje propio, se muestra solo 'Eje *Pz*'.
      - Serv Cont: se muestra combinado como 'Eje *Serv. Cont*' (con punto
        después de 'Serv').
      - Serv Esp: se muestra combinado como 'Eje *Serv Esp*' (sin punto,
        a diferencia de Serv Cont).
    El resto de los ejes conserva el formato normal 'Eje *X* Sub Eje *Y*'.
    """
    if eje == "Pz":
        return f"* Eje *{eje}*"
    if eje == "Serv":
        sub_lower = sub_eje.strip().lower()
        if sub_lower == "cont":
            return "* Eje *Serv. Cont*"
        if sub_lower == "esp":
            return "* Eje *Serv Esp*"
    return f"* Eje *{eje}* Sub Eje *{sub_eje}*"


def procesar_pdf(ruta: str, hora_texto: str, cts_db: dict = None, log_fn=print) -> list:
    """
    Lee el PDF, filtra solo 'Agente de Velocidad',
    queda con el mayor exceso por unidad y genera los mensajes.

    hora_texto: hora que escribió el usuario (24h o 12h am/pm) — define el
    saludo del mensaje (Buenos Días / Buenas Tardes / Buenas Noches).
    cts_db: diccionario {ALIAS: CTS} devuelto por leer_cts_excel(), o None
    si no se adjuntó Excel de CTS (en ese caso, CTS queda en 'S.I.').
    """
    hora, minuto = parsear_hora_24h(hora_texto)
    saludo = saludo_por_hora(hora, minuto)
    registros_velocidad = {}   # alias → {vel, sub_eje, eje}

    with pdfplumber.open(ruta) as pdf:
        total_pags = len(pdf.pages)
        log_fn(f"  Páginas en el PDF: {total_pags}")
        total_filas = 0
        filas_vel   = 0

        for n_pag, page in enumerate(pdf.pages, 1):
            tablas = page.extract_tables()
            for tabla in tablas:
                if not tabla or len(tabla) < 2:
                    continue

                # Detectar fila de encabezados
                header = [str(c).strip().lower() if c else "" for c in tabla[0]]
                if "unidad" not in header:
                    continue

                idx_unidad = next((i for i,h in enumerate(header) if "unidad"    in h), None)
                idx_vel    = next((i for i,h in enumerate(header) if "velocidad" in h), None)
                idx_desc   = next((i for i,h in enumerate(header) if "descripci" in h), None)
                idx_agente = next((i for i,h in enumerate(header) if "agente"    in h), None)

                if idx_unidad is None or idx_vel is None:
                    continue

                for fila in tabla[1:]:
                    if not fila or not fila[idx_unidad]:
                        continue
                    total_filas += 1

                    # Filtrar SOLO alarmas de velocidad (acepta tanto
                    # "Agente de Velocidad" como "Exceso de Velocidad")
                    agente = str(fila[idx_agente]).strip() \
                             if idx_agente is not None and fila[idx_agente] else ""
                    if not es_agente_velocidad(agente):
                        continue

                    filas_vel += 1
                    raw_alias = str(fila[idx_unidad]).strip()
                    alias     = limpiar_alias(raw_alias)

                    # Velocidad de la columna
                    try:
                        vel = float(str(fila[idx_vel]).replace(",",".").strip())
                    except (ValueError, TypeError):
                        vel = 0.0

                    # Descripción
                    desc = str(fila[idx_desc]).strip() if idx_desc is not None \
                           and fila[idx_desc] else ""

                    # Sub-eje y eje desde descripción
                    sub_eje, eje = extraer_subeje_eje(desc)

                    # Excepción manual por unidad (UNIDADES_FORZADAS): ignora
                    # lo que diga el PDF y usa siempre la clasificación fija.
                    forzado = UNIDADES_FORZADAS.get(alias.upper())
                    if forzado:
                        eje, sub_eje = forzado

                    # Guardar solo el mayor exceso por unidad
                    if alias not in registros_velocidad or \
                       vel > registros_velocidad[alias]["vel"]:
                        registros_velocidad[alias] = {
                            "vel":     vel,
                            "sub_eje": sub_eje,
                            "eje":     eje,
                        }

        log_fn(f"  Total filas leídas: {total_filas}")
        log_fn(f"  Filas de velocidad: {filas_vel}")
        log_fn(f"  Unidades únicas con exceso: {len(registros_velocidad)}")

    if not registros_velocidad:
        return []

    # Generar mensajes ordenados por alias
    mensajes = []
    for alias in sorted(registros_velocidad.keys()):
        r       = registros_velocidad[alias]
        vin     = get_vin(alias)
        eje     = r["eje"]     or "S/E"
        sub_eje = r["sub_eje"] or "S/E"
        vel     = r["vel"]
        vel_str = f"{int(vel)} km/h" if vel == int(vel) else f"{vel:.1f} km/h"
        cts     = get_cts(cts_db, alias)
        linea_eje = formatear_linea_eje(eje, sub_eje)

        msg = (
            f"{saludo}\n"
            f"{linea_eje}\n"
            f"* Unidad *{alias}* - Vin *{vin}*\n"
            f"* CTS: *{cts}*\n"
            f"*Incidencia:* Exceso de velocidad\n"
            f"*Incumplimiento:* *{vel_str}* en *autopista*"
        )
        mensajes.append(msg)
        log_fn(f"  ✓ {alias:10s} | {eje:6s} / {sub_eje:20s} | {vel_str}")

    return mensajes


# ── Interfaz gráfica ───────────────────────────────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        root.title("Excesos PDF — MoussaCorp")
        root.geometry("660x600")
        root.configure(bg="#0f1117")
        root.resizable(True, True)

        FM = ("Courier New", 10)
        FB = ("Courier New", 11, "bold")
        BG = "#0f1117"
        FG = "#e2e8f0"

        tk.Label(root, text="EXCESOS DESDE PDF",
                 font=("Courier New", 14, "bold"),
                 bg=BG, fg="#fb923c").pack(pady=(20, 2))
        tk.Label(root,
                 text="Lee PDF de alarmas GTRMax · Agente/Exceso de Velocidad",
                 font=FM, bg=BG, fg="#64748b").pack(pady=(0, 16))

        tk.Button(root,
                  text="📄  Seleccionar PDF de alarmas",
                  font=FB, bg="#7c3aed", fg="white",
                  activebackground="#6d28d9", relief="flat",
                  padx=14, pady=8,
                  command=self.seleccionar_pdf).pack(pady=(0, 4))

        self.lbl_pdf = tk.Label(root, text="Ningún archivo seleccionado",
                                font=FM, bg=BG, fg="#475569")
        self.lbl_pdf.pack(pady=(0, 12))

        tk.Button(root, text="⚡  Generar reportes",
                  font=("Courier New", 12, "bold"),
                  bg="#b45309", fg="white",
                  activebackground="#92400e",
                  relief="flat", padx=20, pady=10,
                  command=self.generar).pack(pady=(0, 10))

        tk.Label(root, text="REPORTES GENERADOS:",
                 font=FB, bg=BG, fg="#94a3b8").pack(anchor="w", padx=20)

        self.txt = scrolledtext.ScrolledText(
            root, font=FM, bg="#0d1117", fg="#fdba74",
            insertbackground=FG, relief="flat", bd=0,
            padx=12, pady=12, wrap="word", height=18, state="disabled"
        )
        self.txt.pack(fill="both", expand=True, padx=20, pady=(4, 10))

        btn_frame = tk.Frame(root, bg=BG)
        btn_frame.pack(pady=(0, 20))

        tk.Button(btn_frame, text="📋  Copiar todo",
                  font=FB, bg="#065f46", fg="white",
                  activebackground="#047857", relief="flat",
                  padx=14, pady=7,
                  command=self.copiar_todo).pack(side="left", padx=6)

        tk.Button(btn_frame, text="💾  Guardar .txt",
                  font=FB, bg="#1e3a5f", fg="white",
                  activebackground="#1e40af", relief="flat",
                  padx=14, pady=7,
                  command=self.guardar_txt).pack(side="left", padx=6)

        self.ruta_pdf = None
        self.mensajes = []

    def _mostrar(self, t):
        self.txt.config(state="normal")
        self.txt.insert("end", t + "\n")
        self.txt.see("end")
        self.txt.config(state="disabled")
        self.root.update()

    def seleccionar_pdf(self):
        r = filedialog.askopenfilename(
            title="Seleccionar PDF de alarmas GTRMax",
            filetypes=[("PDF", "*.pdf")]
        )
        if r:
            self.ruta_pdf = r
            self.lbl_pdf.config(text=os.path.basename(r), fg="#fb923c")

    def generar(self):
        if not self.ruta_pdf:
            messagebox.showinfo("Info", "Primero selecciona el PDF.")
            return

        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.config(state="disabled")

        try:
            self._mostrar(f"Procesando: {os.path.basename(self.ruta_pdf)}\n")
            self.mensajes = procesar_pdf(self.ruta_pdf, log_fn=self._mostrar)

            if not self.mensajes:
                messagebox.showinfo("Sin resultados",
                    "No se encontraron excesos de velocidad en el PDF.\n"
                    "(Solo se procesan filas con Agente = 'Agente de Velocidad' "
                    "o 'Exceso de Velocidad')")
                return

            # Mostrar mensajes separados
            sep = "\n" + "─"*50 + "\n"
            self.txt.config(state="normal")
            self.txt.delete("1.0", "end")
            self.txt.insert("end", sep.join(self.mensajes))
            self.txt.config(state="disabled")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def copiar_todo(self):
        if not self.mensajes:
            messagebox.showinfo("Info", "No hay reportes generados.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append("\n\n".join(self.mensajes))
        messagebox.showinfo("Listo",
            f"✓ {len(self.mensajes)} reporte(s) copiado(s) al portapapeles.")

    def guardar_txt(self):
        if not self.mensajes:
            messagebox.showinfo("Info", "No hay reportes generados.")
            return
        ruta = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt")],
            initialfile=f"excesos_pdf_{datetime.now().strftime('%d%m%Y_%H%M')}.txt"
        )
        if ruta:
            with open(ruta, "w", encoding="utf-8") as f:
                f.write("\n\n".join(self.mensajes))
            messagebox.showinfo("Guardado", f"Guardado en:\n{ruta}")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
