# -*- coding: utf-8 -*-
"""Ejecuta cada herramienta con archivos REALES. Uso:
   python tests/test_paridad.py <carpeta_archivos> <salida.json>
   CP_FORCE_FALLBACK=1 -> fuerza las alternativas de Android (sin calamine/lxml/tkinter)."""
import json, os, sys, tempfile, glob
RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(RAIZ, "app", "src", "main", "python"))
import api

carpeta, destino = sys.argv[1], sys.argv[2]
def f(patron):
    r = glob.glob(os.path.join(carpeta, patron)); return r[0] if r else None

estatus = f("Reporte_de_Estatus_Actual*"); diario = f("Reporte_Operacion_Flota*")
alarmas = f("Detalles_alarmas.pdf"); resumen = f("Resumen_del_Sistema.pdf")
desc = f("Desconexion.xlsx"); disp = f("Disponibilidad.xlsx")
out = tempfile.mkdtemp()
TEXTO = """[9:21 a. m., 11/8/2026] Lon: Buenos Días
* Eje Blv - Sub Eje Acevedo
* Unidad 0645 - Vin 0684
* CTS: Johan S.
Incidencia: Exceso de velocidad
Incumplimiento: 93 km/h en autopista
[9:40 a. m., 11/8/2026] Lon: Buenas Tardes
* Eje Pz
* Unidad 0637 - Vin 0696
* CTS: Wilfredo A.
Incidencia: Exceso de velocidad
Incumplimiento: 101 km/h en autopista
[10:05 a. m., 12/8/2026] Jonaz "Bola 10": Buenos Días
* Eje Serv Esp
* Unidad 0533 - Vin 6307
* CTS: Dorian T.
Incidencia: Exceso de velocidad
Incumplimiento: 102 km/h en autopista
"""
_pre = json.loads(api.ejecutar("alertas", json.dumps(dict(texto=TEXTO, salida=out))))
ALERTAS_XLSX = (_pre.get("archivos") or [None])[0]

casos = {
 "excesos_sin_cts": ("excesos", dict(pdf=alarmas, hora="21:15")),
 "excesos_con_cts": ("excesos", dict(pdf=alarmas, hora="08:02 pm", excel=disp)),
 "sin_info": ("sin_info", dict(archivo=diario)),
 "ultimas": ("ultimas", dict(archivo=estatus, hora="21:15")),
 "ultimas_excluir": ("ultimas", dict(archivo=estatus, hora="21:15", excluir="R-066, 0343")),
 "subeje": ("subeje", dict(archivo=diario, hora="21:00")),
 "alertas": ("alertas", dict(texto=TEXTO, salida=out)),
 "semanal": ("semanal", dict(archivos=[ALERTAS_XLSX, ALERTAS_XLSX], salida=out)),
 "desconexion_sin_status": ("desconexion", dict(pdf=resumen, excel=desc, fecha="01/10/2026 14:30", salida=out)),
 "desconexion_con_status": ("desconexion", dict(pdf=resumen, excel=desc, fecha="01/10/2026 14:30", status_excel=disp, salida=out)),
 "desconexion_nuevo": ("desconexion", dict(pdf=resumen, fecha="01/10/2026", responsable="Prueba", corte="Corte 1", salida=out)),
 "ojo": ("ojo", dict(estatus=estatus, disponibilidad=disp)),
}
res = {"compat": api.ESTADO_COMPAT}
for nombre, (h, p) in casos.items():
    if any(v is None or (isinstance(v, list) and None in v) for v in p.values()):
        res[nombre] = {"omitido": "falta archivo"}; continue
    r = json.loads(api.ejecutar(h, json.dumps(p)))
    r.pop("logs", None); r.pop("detalle", None)
    if r.get("archivos"):                       # comparar contenido, no rutas temporales
        import pandas as pd
        resumen_x = []
        for a in r["archivos"]:
            if a.lower().endswith(".xlsx"):
                xl = pd.ExcelFile(a, engine="openpyxl")
                resumen_x.append({s: xl.parse(s, header=None, dtype=str).fillna("").values.tolist() for s in xl.sheet_names})
        r["archivos"] = resumen_x
    res[nombre] = r
json.dump(res, open(destino, "w"), ensure_ascii=False, indent=1, default=str)
for k, v in res.items():
    if k == "compat": print("compat:", v); continue
    print(f"{k:26} ok={v.get('ok')} msgs={len(v.get('mensajes',[]))} vacio={bool(v.get('vacio'))} err={str(v.get('error',''))[:110]}")
