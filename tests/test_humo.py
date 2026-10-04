"""Pruebas de humo con datos sintéticos (no requieren tus archivos reales).
Para pruebas con archivos reales usa tests/test_paridad.py."""
import json, os, sys, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app", "src", "main", "python"))
import api

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
"""

def run(h, **p):
    return json.loads(api.ejecutar(h, json.dumps(p)))

def test_compat_cargada():
    assert set(api.ESTADO_COMPAT) == {"tkinter", "calamine", "lxml"}

def test_alertas_semanal_mensual():
    out = tempfile.mkdtemp()
    a = run("alertas", texto=TEXTO, salida=out)
    assert a["ok"] and a["total_registros"] == 2
    s = run("semanal", archivos=a["archivos"], salida=out)
    assert s["ok"] and s["archivos"][0].endswith(".xlsx")
    m = run("mensual", archivos=s["archivos"], salida=out)
    assert m["ok"]

def test_errores_entendibles():
    assert not run("alertas", texto="   ")["ok"]
    r = run("ultimas", archivo="/no/existe.xls", hora="21:15")
    assert not r["ok"] and r["error"]
    assert "desconocida" in run("inventada")["error"]

def test_exclusion_unidades():
    v, i = api.parsear_lista_exclusion("R-66, 0343 y xx")
    assert "R-066" in v and "0343" in v and i == ["xx"]
