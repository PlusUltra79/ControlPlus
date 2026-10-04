# -*- coding: utf-8 -*-
"""
cp_compat.py — Capa de compatibilidad Android (MoussaCorp · ControlPlus CCO)
============================================================================
Permite ejecutar las herramientas originales (tools/*.py) SIN modificar su
lógica de negocio dentro de la APK:

  1. tkinter no existe en Android -> se instalan módulos "fantasma" (solo si
     tkinter real no está disponible). Las clases de GUI nunca se usan.
  2. python-calamine (Rust) puede no estar disponible -> pd.ExcelFile /
     pd.read_excel con engine="calamine" caen a openpyxl, reparando antes el
     stylesheet XML inválido que trae el Excel real de "Disponibilidad".
  3. lxml puede no estar disponible -> "lxml" cae a "html5lib", que corrige
     las etiquetas sin cerrar igual que un navegador.

Si las librerías reales SÍ están disponibles (PC de escritorio), no se toca nada.
"""
import os
import re
import sys
import types
import zipfile
import tempfile

FORZAR_FALLBACK = os.environ.get("CP_FORCE_FALLBACK") == "1"   # solo para pruebas
ESTADO = {"tkinter": "real", "calamine": "real", "lxml": "real"}


# ── 1) tkinter ───────────────────────────────────────────────────────────────
class _ModuloFantasma(types.ModuleType):
    def __getattr__(self, nombre):
        if nombre.startswith("__"):
            raise AttributeError(nombre)
        clase = type(nombre, (), {
            "__init__": lambda self, *a, **k: None,
            "__getattr__": lambda self, n: (lambda *a, **k: None),
        })
        setattr(self, nombre, clase)
        return clase


def _instalar_tkinter_fantasma():
    for nombre in ("tkinter", "tkinter.filedialog", "tkinter.messagebox",
                   "tkinter.scrolledtext", "tkinter.ttk"):
        sys.modules[nombre] = _ModuloFantasma(nombre)
    for sub in ("filedialog", "messagebox", "scrolledtext", "ttk"):
        setattr(sys.modules["tkinter"], sub, sys.modules["tkinter." + sub])
    ESTADO["tkinter"] = "fantasma"


def _preparar_tkinter():
    if FORZAR_FALLBACK:
        _instalar_tkinter_fantasma()
        return
    try:
        import tkinter  # noqa: F401
    except Exception:
        _instalar_tkinter_fantasma()


# ── 2) calamine -> openpyxl (con reparación de styles.xml) ───────────────────
_STYLES_MINIMO = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    '</styleSheet>'
)


def _reparar_xlsx(ruta):
    """Copia el .xlsx a un temporal con xl/styles.xml mínimo y válido. Los datos
    no dependen del stylesheet; los ceros a la izquierda se conservan porque
    viven como texto en sharedStrings."""
    destino = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
    with zipfile.ZipFile(ruta) as zin, zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            datos = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                datos = _STYLES_MINIMO.encode("utf-8")
            zout.writestr(item, datos)
    return destino


def _calamine_disponible():
    if FORZAR_FALLBACK:
        return False
    try:
        import python_calamine  # noqa: F401
        return True
    except Exception:
        return False


def _instalar_shim_excel():
    import pandas as pd
    if _calamine_disponible():
        return
    ESTADO["calamine"] = "openpyxl+reparación"
    _ExcelFile_orig = pd.ExcelFile
    _read_excel_orig = pd.read_excel

    def _ruta_openpyxl(ruta):
        if isinstance(ruta, (str, os.PathLike)) and str(ruta).lower().endswith((".xlsx", ".xlsm")):
            return _reparar_xlsx(str(ruta))
        return ruta

    def ExcelFile(path_or_buffer, engine=None, *a, **k):
        if engine == "calamine":
            engine = "openpyxl"
            path_or_buffer = _ruta_openpyxl(path_or_buffer)
        return _ExcelFile_orig(path_or_buffer, engine, *a, **k)

    def read_excel(io_, *a, **k):
        if k.get("engine") == "calamine":
            k["engine"] = "openpyxl"
            io_ = _ruta_openpyxl(io_)
        return _read_excel_orig(io_, *a, **k)

    pd.ExcelFile = ExcelFile
    pd.read_excel = read_excel


# ── 3) lxml -> normalizador propio de tablas HTML ───────────────────────────
# Los reportes del GTRMax traen <tr>/<td> sin cerrar. lxml los auto-cierra; en
# Android, html5lib reubica mal la tabla anidada (probado con el archivo real),
# así que se normaliza el HTML cerrando filas/celdas por tabla y luego se usa el
# parser estándar 'html.parser', que ya recibe un HTML bien formado.
from html.parser import HTMLParser


class _NormalizadorTablas(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=False)
        self.salida = []
        self.pila = []                       # una entrada por <table>: {"fila": bool, "celda": nombre|None}

    def _cerrar_celda(self):
        if self.pila and self.pila[-1]["celda"]:
            self.salida.append("</%s>" % self.pila[-1]["celda"])
            self.pila[-1]["celda"] = None

    def _cerrar_fila(self):
        self._cerrar_celda()
        if self.pila and self.pila[-1]["fila"]:
            self.salida.append("</tr>")
            self.pila[-1]["fila"] = False

    def handle_starttag(self, tag, attrs):
        crudo = self.get_starttag_text()
        if tag == "table":
            self.salida.append(crudo)
            self.pila.append({"fila": False, "celda": None})
        elif tag in ("thead", "tbody", "tfoot") and self.pila:
            self._cerrar_fila()
            self.salida.append(crudo)
        elif tag == "tr" and self.pila:
            self._cerrar_fila()
            self.salida.append(crudo)
            self.pila[-1]["fila"] = True
        elif tag in ("td", "th") and self.pila:
            self._cerrar_celda()          # una celda suelta NO crea <tr> (como libxml2)
            self.salida.append(crudo)
            self.pila[-1]["celda"] = tag
        else:
            self.salida.append(crudo)

    def handle_startendtag(self, tag, attrs):
        self.salida.append(self.get_starttag_text())

    def handle_endtag(self, tag):
        if tag == "table" and self.pila:
            self._cerrar_fila()
            self.salida.append("</table>")
            self.pila.pop()
        elif tag in ("thead", "tbody", "tfoot") and self.pila:
            self._cerrar_fila()
            self.salida.append("</%s>" % tag)
        elif tag == "tr" and self.pila:
            self._cerrar_fila()
        elif tag in ("td", "th") and self.pila:
            self._cerrar_celda()
        else:
            self.salida.append("</%s>" % tag)

    def handle_data(self, data):
        self.salida.append(data)

    def handle_entityref(self, name):
        self.salida.append("&%s;" % name)

    def handle_charref(self, name):
        self.salida.append("&#%s;" % name)

    def handle_comment(self, data):
        pass

    def close(self):
        super().close()
        while self.pila:
            self._cerrar_fila()
            self.salida.append("</table>")
            self.pila.pop()


def normalizar_html(texto):
    n = _NormalizadorTablas()
    n.feed(texto)
    n.close()
    return "".join(n.salida)


def _leer_texto(ruta):
    with open(ruta, "rb") as f:
        crudo = f.read()
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return crudo.decode(enc)
        except UnicodeDecodeError:
            continue
    return crudo.decode("latin-1", errors="replace")


def _instalar_shim_html():
    try:
        if FORZAR_FALLBACK:
            raise ImportError
        import lxml  # noqa: F401
        return
    except Exception:
        pass
    ESTADO["lxml"] = "normalizador-html"
    import io as _io
    import pandas as pd
    import bs4

    _BS_orig = bs4.BeautifulSoup

    class BeautifulSoup(_BS_orig):
        def __init__(self, markup="", features=None, *a, **k):
            if features == "lxml":
                features = "html.parser"
                if isinstance(markup, bytes):
                    markup = markup.decode("utf-8", errors="replace")
                if isinstance(markup, str):
                    markup = normalizar_html(markup)
            super().__init__(markup, features, *a, **k)

    bs4.BeautifulSoup = BeautifulSoup   # debe ejecutarse ANTES de importar tools

    _read_html_orig = pd.read_html

    def _tablas_como_pandas(texto):
        """Réplica de pd.read_html(flavor='lxml') para reportes GTRMax: filas = <tr>
        hijas DIRECTAS de la tabla (o de thead/tbody/tfoot); mismo motor de
        inferencia de tipos de pandas (TextParser) para obtener el mismo DataFrame."""
        from pandas.io.parsers import TextParser
        soup = _BS_orig(normalizar_html(texto), "html.parser")
        resultado = []
        for tabla in soup.find_all("table"):
            filas = []
            for hijo in tabla.find_all(["tr", "thead", "tbody", "tfoot"], recursive=False):
                trs = [hijo] if hijo.name == "tr" else hijo.find_all("tr", recursive=False)
                for tr in trs:
                    celdas = [" ".join(c.get_text().split()) for c in tr.find_all(["td", "th"], recursive=False)]
                    if celdas:
                        filas.append(celdas)
            if filas:
                ancho = max(len(f) for f in filas)
                filas = [f + [""] * (ancho - len(f)) for f in filas]
                resultado.append(TextParser(filas, header=None, thousands=",").read())
        if not resultado:
            raise ValueError("No tables found")
        return resultado

    def read_html(io_, *a, **k):
        if k.get("flavor") == "lxml":
            if isinstance(io_, (str, os.PathLike)) and os.path.exists(str(io_)):
                return _tablas_como_pandas(_leer_texto(str(io_)))
            k["flavor"] = "html5lib"
        return _read_html_orig(io_, *a, **k)

    pd.read_html = read_html


def preparar():
    _preparar_tkinter()
    _instalar_shim_excel()
    _instalar_shim_html()
    return dict(ESTADO)
