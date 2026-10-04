"""
UT en Movimiento + Sin Información por Sub-Eje
--------------------------------------------------
Lee el reporte diario GTR (.xls HTML disfrazado) y genera DOS tipos
de reporte a partir de la misma regla de "en movimiento":

  Regla: solo unidades con Distancia Recorrida (Km) > 1.0

1) REPORTE POR EJE (idéntico a ut_en_movimiento.py)
   Un solo mensaje combinado, con una sección por Eje y un Total
   General al final.

2) REPORTE "SIN INFORMACIÓN" POR SUB-EJE (nuevo)
   Un mensaje INDEPENDIENTE por cada Eje, donde dentro de ese mensaje
   cada Sub-eje aparece como su propia sub-sección con numeración que
   reinicia en 1, y un Total al final sumando todos los sub-ejes de
   ese Eje.

Formato del reporte por Eje:
  *UT en movimiento para la hora: 08:00*

  *Am*
  1. R-003
  2. 0491
  * Total: 13

  *Blv*
  ...
  * Total General: Y

Formato del reporte Sin Información por Sub-eje (uno de estos por Eje):
  UT en movimiento para la hora: 21:00
  Sin información
  Páez Andres Bello
  1. 0648
  2. 0636
  Pedro Gual
  1. R-034
  2. 0477
  * Total: 4

Compatible con archivos .xls exportados por el sistema GTR.
Requiere: pip install beautifulsoup4 lxml

Uso: python ut_movimiento_subeje.py
"""

from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os

# ── Configuración ─────────────────────────────────────────────────────────────
ORDEN_FLOTAS = ["AM", "BLV", "MET", "OCM", "PZ", "SERV."]

NOMBRE_FLOTA = {
    "AM":    "Am",
    "BLV":   "Blv",
    "MET":   "Met",
    "OCM":   "Ocm",
    "PZ":    "Pz",
    "SERV.": "Serv",
}

# Regla de "en movimiento": misma que ut_en_movimiento.py
UMBRAL_DISTANCIA_KM = 1.0


# ── Lectura del archivo ───────────────────────────────────────────────────────

def leer_archivo(ruta):
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            with open(ruta, encoding=enc) as f:
                contenido = f.read()
            break
        except UnicodeDecodeError:
            continue

    soup = BeautifulSoup(contenido, "html.parser")
    tabla = soup.find("table")
    if not tabla:
        raise ValueError("No se encontró ninguna tabla en el archivo.")

    filas = tabla.find_all("tr")
    encabezados = None
    datos = []
    for fila in filas:
        celdas = fila.find_all(["th", "td"])
        textos = [c.get_text(strip=True) for c in celdas]
        if encabezados is None:
            if "Flota" in textos and "Alias" in textos:
                encabezados = textos
        else:
            if len(textos) == len(encabezados):
                datos.append(textos)

    if encabezados is None:
        raise ValueError("No se encontró la fila de encabezados.")

    return pd.DataFrame(datos, columns=encabezados)


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


def limpiar_alias(raw: str) -> str:
    """Quita el prefijo 'UT:' que a veces viene del archivo GTR."""
    return re.sub(r"(?i)^ut:\s*", "", str(raw)).strip()


def _preparar_df(df: pd.DataFrame):
    """Detecta columnas y devuelve (df_movimiento, col_flota, col_subflota, col_alias)."""
    columnas = list(df.columns)
    col_flota    = encontrar_columna(columnas, ["Flota"])
    col_subflota = encontrar_columna(columnas, ["Sub Flota", "Sub-Flota", "SubFlota"])
    col_alias    = encontrar_columna(columnas, ["Alias"])
    col_dist     = encontrar_columna(columnas, [
        "Distancia Recorrida (Km)", "Distancia Recorrida",
        "distancia recorrida (km)", "Distancia"
    ])

    faltantes = [n for n, c in [
        ("Flota", col_flota), ("Sub Flota", col_subflota),
        ("Alias", col_alias), ("Distancia Recorrida (Km)", col_dist)
    ] if not c]
    if faltantes:
        raise ValueError(
            f"Columnas no encontradas: {', '.join(faltantes)}\n"
            f"Columnas detectadas: {', '.join(columnas)}"
        )

    df = df.copy()
    df["_flota_up"]    = df[col_flota].astype(str).str.strip().str.upper()
    df["_subflota_up"] = df[col_subflota].astype(str).str.strip().str.upper()
    df[col_dist] = pd.to_numeric(
        df[col_dist].astype(str).str.replace(",", ".", regex=False).str.strip(),
        errors="coerce"
    ).fillna(0)

    df_mov = df[
        df["_flota_up"].isin(set(ORDEN_FLOTAS)) & (df[col_dist] > UMBRAL_DISTANCIA_KM)
    ].copy()

    return df_mov, col_flota, col_subflota, col_alias


# ── Reporte 1: por Eje (idéntico a ut_en_movimiento.py) ───────────────────────

def generar_reporte_eje(df: pd.DataFrame, hora: str) -> str:
    """
    Reporte combinado por Eje.
    Retorna: str, mensaje único con todas las secciones + Total General.
    """
    df_mov, col_flota, col_subflota, col_alias = _preparar_df(df)

    secciones = []
    total_general = 0

    for flota_key in ORDEN_FLOTAS:
        nombre_eje = NOMBRE_FLOTA.get(flota_key, flota_key)
        grupo = df_mov[df_mov["_flota_up"] == flota_key]

        lineas = [f"*{nombre_eje}*"]
        count = 0
        for _, fila in grupo.iterrows():
            alias = limpiar_alias(fila[col_alias])
            count += 1
            lineas.append(f"{count}. {alias}")
        lineas.append(f"* Total: {count}")

        secciones.append("\n".join(lineas))
        total_general += count

    cuerpo = "\n\n".join(secciones)
    header = f"*UT en movimiento para la hora: {hora}*"
    return f"{header}\n\n{cuerpo}\n\n* Total General: {total_general}"


# ── Reporte 2: Sin Información por Sub-Eje (nuevo) ────────────────────────────

def generar_reportes_subeje(df: pd.DataFrame, hora: str) -> list:
    """
    Reporte "Sin información" con desglose por Sub-eje.
    Retorna: list[tuple]: [(nombre_eje, mensaje), ...]
    Un mensaje independiente por cada Eje; dentro de cada mensaje,
    cada Sub-eje es su propia sub-sección con numeración reiniciada en 1.
    """
    df_mov, col_flota, col_subflota, col_alias = _preparar_df(df)

    resultados = []

    for flota_key in ORDEN_FLOTAS:
        nombre_eje = NOMBRE_FLOTA.get(flota_key, flota_key)
        grupo_eje = df_mov[df_mov["_flota_up"] == flota_key]

        # Sub-ejes presentes en este eje, orden alfabético para salida consistente
        subejes_presentes = sorted(grupo_eje["_subflota_up"].unique().tolist())

        bloques = []
        total_eje = 0
        for subeje_up in subejes_presentes:
            grupo_sub = grupo_eje[grupo_eje["_subflota_up"] == subeje_up]

            subeje_raw = str(grupo_sub.iloc[0][col_subflota]).strip()
            subeje_display = subeje_raw.title() if subeje_raw else "S/E"

            lineas = [subeje_display]
            count = 0
            for _, fila in grupo_sub.iterrows():
                alias = limpiar_alias(fila[col_alias])
                count += 1
                lineas.append(f"{count}. {alias}")

            bloques.append("\n".join(lineas))
            total_eje += count

        header = f"UT en movimiento para la hora: {hora}\nSin información"
        if bloques:
            cuerpo = "\n".join(bloques)
            mensaje = f"{header}\n{cuerpo}\n* Total: {total_eje}"
        else:
            mensaje = f"{header}\n* Total: {total_eje}"

        resultados.append((nombre_eje, mensaje))

    return resultados


def formatear_salida_completa(reporte_eje: str, reportes_subeje: list) -> str:
    partes = [
        f"{'═'*50}\n  REPORTE POR EJE\n{'═'*50}\n{reporte_eje}",
    ]
    for nombre_eje, mensaje in reportes_subeje:
        partes.append(
            f"{'═'*50}\n  SIN INFORMACIÓN — EJE {nombre_eje.upper()}\n{'═'*50}\n{mensaje}"
        )
    return "\n\n".join(partes)


# ── Interfaz gráfica ──────────────────────────────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        root.title("UT en Movimiento + Sin Información por Sub-Eje")
        root.geometry("680x680")
        root.configure(bg="#0f1117")
        root.resizable(True, True)

        FONT_MONO = ("Courier New", 10)
        FONT_BOLD = ("Courier New", 11, "bold")
        BG     = "#0f1117"
        FG     = "#e2e8f0"
        ACCENT = "#3fb950"
        BTN_BG = "#1a7431"

        tk.Label(root, text="UT EN MOVIMIENTO · SIN INFO POR SUB-EJE",
                 font=("Courier New", 13, "bold"),
                 bg=BG, fg=ACCENT).pack(pady=(20, 2))
        tk.Label(root, text="Reporte por Eje + desglose Sin Información por Sub-eje",
                 font=FONT_MONO, bg=BG, fg="#64748b").pack(pady=(0, 14))

        # ── Hora manual ──
        frame_hora = tk.Frame(root, bg=BG)
        frame_hora.pack(pady=(0, 10))
        tk.Label(frame_hora, text="Hora del reporte (ej: 21:00):",
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

        tk.Button(root, text="⚡  Generar reportes",
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
            padx=10, pady=10, wrap="word", height=20
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
        self.reporte_eje = ""
        self.reportes_subeje = []

    def cargar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo GTR",
            filetypes=[("Archivos GTR", "*.xls *.xlsx")]
        )
        if not ruta:
            return
        self.ruta_archivo = ruta
        self.lbl_archivo.config(text=os.path.basename(ruta), fg="#3fb950")

    def generar(self):
        if not self.ruta_archivo:
            messagebox.showinfo("Info", "Primero selecciona el archivo GTR.")
            return

        hora = self.entry_hora.get().strip()
        if not hora:
            messagebox.showinfo("Info", "Ingresa la hora del reporte (ej: 21:00).")
            return

        try:
            df = leer_archivo(self.ruta_archivo)
            self.reporte_eje = generar_reporte_eje(df, hora)
            self.reportes_subeje = generar_reportes_subeje(df, hora)
            texto = formatear_salida_completa(self.reporte_eje, self.reportes_subeje)
            self.txt.delete("1.0", "end")
            self.txt.insert("end", texto)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def copiar_todo(self):
        if not self.reporte_eje and not self.reportes_subeje:
            messagebox.showinfo("Info", "No hay reportes generados.")
            return
        partes = [self.reporte_eje] + [m for _, m in self.reportes_subeje]
        self.root.clipboard_clear()
        self.root.clipboard_append("\n\n".join(partes))
        messagebox.showinfo("Listo", "Reportes copiados al portapapeles.")

    def guardar_txt(self):
        if not self.reporte_eje and not self.reportes_subeje:
            messagebox.showinfo("Info", "No hay reportes generados.")
            return
        ruta = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Archivo de texto", "*.txt")],
            initialfile=f"ut_movimiento_subeje_{datetime.now().strftime('%d%m%Y_%H%M')}.txt"
        )
        if not ruta:
            return
        with open(ruta, "w", encoding="utf-8") as f:
            f.write(formatear_salida_completa(self.reporte_eje, self.reportes_subeje))
        messagebox.showinfo("Guardado", f"Archivo guardado en:\n{ruta}")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
