# -*- coding: utf-8 -*-
"""
api.py — Fachada Python de ControlPlus CCO (reemplaza a bot.py)
================================================================
La interfaz de la APK llama a ejecutar(herramienta, json) y recibe un JSON:

  {"ok": true,  "mensajes": [{"titulo": "...", "texto": "..."}],
                "archivos": ["/ruta/Consolidado.xlsx"], "logs": ["..."]}
  {"ok": false, "error": "Mensaje entendible para el usuario"}

No hay red, ni Telegram, ni token: todo se procesa en el teléfono.
La lógica de negocio sigue viviendo, intacta, en tools/*.py.
"""
import json
import os
import re
import traceback

import cp_compat
ESTADO_COMPAT = cp_compat.preparar()          # SIEMPRE antes de importar tools

from tools import excesos_pdf as m_excesos                      # noqa: E402
from tools import sin_informacion as m_sin_info                 # noqa: E402
from tools import ultimas_ut_movimiento as m_ultimas            # noqa: E402
from tools import ut_movimiento_subeje as m_subeje              # noqa: E402
from tools import consolidador_semanal as m_semanal             # noqa: E402
from tools import consolidador_mensual as m_mensual             # noqa: E402
from tools import actualizar_desconexion_V2_ as m_desc          # noqa: E402
from tools import data_alertas as m_alertas                     # noqa: E402
from tools import ojo_de_halcon as m_ojo                        # noqa: E402
from tools import resumen_recorrido as m_recorrido              # noqa: E402

AQUI = os.path.dirname(os.path.abspath(__file__))


def _resolver_rutas_db():
    """Dentro de la APK (Chaquopy) los .py/.json viajan empaquetados y open() sobre
    su ruta puede fallar. Si el archivo no existe en disco, se lee con el loader
    de Python y se copia a una carpeta temporal real."""
    ruta = os.path.join(AQUI, "rutas_db.json")
    if os.path.exists(ruta):
        return ruta
    import tempfile
    datos = __loader__.get_data(ruta)
    destino = os.path.join(tempfile.gettempdir(), "rutas_db.json")
    with open(destino, "wb") as f:
        f.write(datos)
    return destino


RUTAS_DB = _resolver_rutas_db()
VERSION = "1.0.0"


# ── Utilidades (portadas de bot.py) ──────────────────────────────────────────
def parsear_lista_exclusion(texto):
    """'R-066, 0343 y R-7' -> ({'R-066','0343','R-007'}, [no reconocidos])"""
    tokens = [t for t in re.split(r"[,\n;]+|\s+y\s+", texto or "") if t.strip()]
    validas, invalidas = set(), []
    for t in tokens:
        alias = m_recorrido.normalizar_unidad(t)
        if alias:
            validas.add(alias)
        else:
            invalidas.append(t.strip())
    return validas, invalidas


def _excluir(params, logs):
    texto = (params.get("excluir") or "").strip()
    if not texto:
        return None
    validas, invalidas = parsear_lista_exclusion(texto)
    if invalidas:
        logs.append("No reconocí: " + ", ".join(invalidas) + " (se ignoran).")
    if not validas:
        raise ValueError("No reconocí ninguna unidad para obviar. Escríbelas como R-036 o 0140, separadas por coma.")
    logs.append("Obviando: " + ", ".join(sorted(validas)))
    return validas


def _req(params, *claves):
    for c in claves:
        if not params.get(c):
            raise ValueError("Falta un dato obligatorio: " + c)


def _msgs(lista, titulo_fn=None):
    return [{"titulo": titulo_fn(i, m) if titulo_fn else "", "texto": m} for i, m in enumerate(lista)]


def _salida(params):
    d = params.get("salida") or os.path.join(AQUI, "salida")
    os.makedirs(d, exist_ok=True)
    return d


# ── Herramientas ─────────────────────────────────────────────────────────────
def _t_excesos(p, logs):
    _req(p, "pdf", "hora")
    m_excesos.parsear_hora_24h(p["hora"])
    cts_db = m_excesos.leer_cts_excel(p["excel"], logs.append) if p.get("excel") else None
    mensajes = m_excesos.procesar_pdf(p["pdf"], p["hora"], cts_db, logs.append)
    if not mensajes:
        return {"vacio": "No se encontraron excesos de velocidad en el PDF "
                         "(solo se procesan filas con Agente de Velocidad o Exceso de Velocidad)."}
    return {"mensajes": _msgs(mensajes, lambda i, m: f"Exceso {i + 1}")}


def _t_sin_info(p, logs):
    _req(p, "archivo")
    df = m_sin_info.leer_archivo(p["archivo"])
    todos = m_sin_info.generar_mensajes(df)
    out = []
    for eje, msgs in todos:
        if msgs:
            out.append({"titulo": f"Eje {eje} · {len(msgs)} unidad(es)",
                        "texto": f"— Eje {eje} — ({len(msgs)} unidad(es)) —\n\n" + "\n\n".join(msgs)})
    if not out:
        return {"vacio": "No se encontraron unidades para reportar."}
    return {"mensajes": out}


def _t_ultimas(p, logs):
    _req(p, "archivo", "hora")
    m_ultimas.parsear_hora_manual(p["hora"])
    excluir = _excluir(p, logs)
    df = m_ultimas.leer_archivo(p["archivo"])
    res = m_ultimas.generar_reporte(df, p["hora"], logs.append, excluir)
    return {"mensajes": [{"titulo": f"Grupo {n}", "texto": m} for n, m in res]}


def _t_subeje(p, logs):
    _req(p, "archivo", "hora")
    df = m_subeje.leer_archivo(p["archivo"])
    eje = m_subeje.generar_reporte_eje(df, p["hora"])
    subs = m_subeje.generar_reportes_subeje(df, p["hora"])
    out = [{"titulo": "Reporte por Eje", "texto": eje}]
    out += [{"titulo": f"Sin información · {n}", "texto": m} for n, m in subs]
    return {"mensajes": out}


def _consolidar(modulo, fn_leer, p, logs):
    archivos = p.get("archivos") or []
    if not archivos:
        raise ValueError("Selecciona al menos un archivo Excel.")
    df = getattr(modulo, fn_leer)(archivos, logs.append)
    if df is None:
        raise ValueError("No se pudo generar el consolidado. Revisa los archivos seleccionados.")
    ruta = modulo.guardar_consolidado(df, _salida(p), logs.append)
    return {"archivos": [ruta] if ruta else []}


def _t_semanal(p, logs):
    return _consolidar(m_semanal, "consolidar_archivos", p, logs)


def _t_mensual(p, logs):
    return _consolidar(m_mensual, "consolidar_semanales", p, logs)


def _t_desconexion(p, logs):
    _req(p, "pdf", "fecha")
    fecha = m_desc._parsear_fecha_actual(p["fecha"].strip())
    status_db = m_desc.leer_status_excel(p["status_excel"], logs.append) if p.get("status_excel") else None
    ruta = m_desc.actualizar_desconexion(
        p["pdf"], p.get("excel") or None, fecha, _salida(p), logs.append,
        p.get("responsable", "") or "", p.get("corte", "") or "", status_db)
    return {"archivos": [ruta]}


def _t_alertas(p, logs):
    texto = m_alertas._limpiar_texto_whatsapp(p.get("texto") or "")
    if not texto.strip():
        raise ValueError("Pega primero el texto de los reportes de alertas.")
    records = m_alertas.parse_reports(texto)
    if not records:
        raise ValueError("No reconocí ningún reporte en el texto pegado. Revisa que incluya la hora y fecha de cada mensaje.")
    logs.append(f"{len(records)} registro(s) reconocido(s).")
    from datetime import datetime
    salida = os.path.join(_salida(p), f"Data_Alertas_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx")
    ruta = m_alertas.build_excel(records, p.get("excel") or None, salida)
    return {"archivos": [ruta], "total_registros": len(records)}


def _t_ojo(p, logs):
    _req(p, "estatus", "disponibilidad")
    excluir = _excluir(p, logs)
    msgs = m_ojo.generar_reportes_ojo_de_halcon(
        p["estatus"], p["disponibilidad"], RUTAS_DB, None, logs.append, excluir)
    if not msgs:
        return {"vacio": "Sin incidencias detectadas en este reporte. ✅"}
    return {"mensajes": _msgs(msgs, lambda i, m: f"Incidencia {i + 1}")}


def _t_recorrido(p, logs):
    _req(p, "archivo", "unidad", "cts", "tipo")
    alias = m_recorrido.normalizar_unidad(p["unidad"])
    if not alias:
        raise ValueError("No reconozco esa unidad. Escríbela como R-066 o 0343.")
    info = m_recorrido.buscar_unidad(alias)
    if info is None:
        info = m_recorrido.parsear_unidad_manual(p.get("manual") or "")
        if not info:
            raise ValueError(f"La unidad {alias} no está en la lista. Escribe: Eje, Sub eje, VIN (ej. Blv, Acevedo, 1234).")
    msgs = m_recorrido.generar_resumen(p["archivo"], alias, p["cts"], p["tipo"], RUTAS_DB, info, logs.append)
    return {"mensajes": _msgs(msgs, lambda i, m: f"Resumen {i + 1}")}


HERRAMIENTAS = {
    "excesos": _t_excesos, "sin_info": _t_sin_info, "ultimas": _t_ultimas,
    "subeje": _t_subeje, "semanal": _t_semanal, "mensual": _t_mensual,
    "desconexion": _t_desconexion, "alertas": _t_alertas, "ojo": _t_ojo,
    "recorrido": _t_recorrido,
}


# ── Puntos de entrada para Kotlin / JS ───────────────────────────────────────
def ejecutar(herramienta, params_json):
    logs = []
    try:
        params = json.loads(params_json or "{}")
        fn = HERRAMIENTAS.get(herramienta)
        if fn is None:
            raise ValueError("Herramienta desconocida: " + str(herramienta))
        r = fn(params, logs)
        r.update({"ok": True, "logs": [str(x) for x in logs]})
        return json.dumps(r, ensure_ascii=False)
    except ValueError as e:                         # mensajes pensados para el usuario
        return json.dumps({"ok": False, "error": str(e), "logs": logs}, ensure_ascii=False)
    except Exception as e:                          # error inesperado: detalle solo en el log
        return json.dumps({"ok": False,
                           "error": "No se pudo procesar el archivo. Verifica que sea el reporte correcto "
                                    "y vuelve a intentar. (" + type(e).__name__ + ": " + str(e)[:160] + ")",
                           "detalle": traceback.format_exc()[-1500:], "logs": logs}, ensure_ascii=False)


def estado():
    return json.dumps({"version": VERSION, "compat": ESTADO_COMPAT}, ensure_ascii=False)


def unidad_conocida(unidad):
    """Para la UI de Recorrido: ¿la unidad está en la lista de ejes/VIN?"""
    alias = m_recorrido.normalizar_unidad(unidad or "")
    if not alias:
        return json.dumps({"alias": None})
    info = m_recorrido.buscar_unidad(alias)
    return json.dumps({"alias": alias, "info": list(info) if info else None}, ensure_ascii=False)
