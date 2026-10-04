"""
Consolidador Semanal de Reportes
---------------------------------
Lee varios archivos Excel diarios y genera un único archivo consolidado semanal,
con la misma estructura de columnas del reporte semanal oficial.

Estructura diario  → Estructura semanal (hoja "General"):
  Número de reporte → Reporte
  Sub‑eje           → Sub Eje
  Respuestas        → Respuesta

Requiere: pip install pandas openpyxl
          tkinter viene incluido con Python en Windows

Uso: python consolidador_semanal.py
"""

import os, re, glob
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# ── Configuración ──────────────────────────────────────────────────────────────
NOMBRE_HOJA_ENTRADA = None   # None = primera hoja; cambia si tu diario usa otro nombre

# Columnas del archivo SEMANAL en el orden exacto
COLUMNAS_SEMANAL = [
    "Responsable", "Reporte", "Fecha", "Hora",
    "Eje", "Sub Eje", "Unidad", "Cts",
    "Tipo de reporte", "Incidencias", "Niveles",
    "Descripción", "Observación", "Respuesta", "Detalle"
]

# Mapeo de nombres del diario → nombres del semanal
RENOMBRAR = {
    "Número de reporte": "Reporte",
    "Sub‑eje":           "Sub Eje",
    "Sub-eje":           "Sub Eje",   # por si viene con guion normal
    "Respuestas":        "Respuesta",
}

NOMBRE_HOJA_SALIDA = "General"
GENERAR_CSV        = True

# ── Funciones ──────────────────────────────────────────────────────────────────

def inferir_fecha_desde_nombre(nombre_archivo: str):
    """Extrae fecha YYYY-MM-DD del nombre del archivo si existe."""
    m = re.search(r"(\d{4}-\d{2}-\d{2})", nombre_archivo)
    if m:
        return m.group(1)
    m2 = re.search(r"(\d{2})[-/](\d{2})[-/](\d{4})", nombre_archivo)
    if m2:
        d, mo, y = m2.groups()
        return f"{y}-{mo}-{d}"
    return None


def leer_archivo_diario(ruta_archivo: str, log_fn=print):
    """Lee un archivo Excel diario, renombra columnas y retorna DataFrame normalizado."""
    nombre = os.path.basename(ruta_archivo)
    try:
        kwargs = {"engine": "openpyxl"}
        if NOMBRE_HOJA_ENTRADA:
            kwargs["sheet_name"] = NOMBRE_HOJA_ENTRADA

        df = pd.read_excel(ruta_archivo, **kwargs)

        if df.empty:
            log_fn(f"  ⚠ {nombre}: archivo vacío, se omite.")
            return None

        # Renombrar columnas al estándar semanal
        df.rename(columns=RENOMBRAR, inplace=True)

        # Si no existe columna Fecha, inferir desde nombre
        if "Fecha" not in df.columns:
            fecha = inferir_fecha_desde_nombre(nombre)
            df["Fecha"] = fecha if fecha else None
            if fecha:
                log_fn(f"  ℹ {nombre}: Fecha inferida desde nombre → {fecha}")
            else:
                log_fn(f"  ⚠ {nombre}: no se pudo inferir Fecha.")

        log_fn(f"  ✓ {nombre}: {len(df)} filas leídas.")
        return df

    except Exception as e:
        log_fn(f"  ✗ {nombre}: ERROR → {e}")
        return None


def consolidar_archivos(lista_rutas: list, log_fn=print):
    """Concatena todos los diarios, limpia y ordena."""
    dataframes = []

    for ruta in lista_rutas:
        df = leer_archivo_diario(ruta, log_fn)
        if df is not None:
            dataframes.append(df)

    if not dataframes:
        log_fn("✗ No se pudo leer ningún archivo.")
        return None

    log_fn(f"\nUniendo {len(dataframes)} archivo(s)...")
    consolidado = pd.concat(dataframes, ignore_index=True, sort=False)

    # Eliminar filas completamente vacías
    consolidado.dropna(how="all", inplace=True)

    # Eliminar duplicados exactos
    antes = len(consolidado)
    consolidado.drop_duplicates(keep="first", inplace=True)
    dup = antes - len(consolidado)
    if dup:
        log_fn(f"  ℹ {dup} fila(s) duplicada(s) eliminada(s).")

    # Ordenar por Fecha (conservando el orden original dentro de cada fecha)
    # Se agrega un índice temporal para respetar el orden de entrada dentro del mismo día
    consolidado["_orden_original"] = range(len(consolidado))

    if "Fecha" in consolidado.columns:
        consolidado["_fecha_dt"] = pd.to_datetime(
            consolidado["Fecha"], dayfirst=True, errors="coerce"
        )
        consolidado.sort_values(
            by=["_fecha_dt", "_orden_original"],
            inplace=True, ignore_index=True
        )
        consolidado.drop(columns=["_fecha_dt"], inplace=True)

    consolidado.drop(columns=["_orden_original"], inplace=True)

    # Renumerar columna Reporte de forma consecutiva global (N°001, N°002, ...)
    if "Reporte" in consolidado.columns:
        consolidado["Reporte"] = [f"N°{i:03d}" for i in range(1, len(consolidado) + 1)]
        log_fn(f"  ℹ Columna 'Reporte' renumerada del N°001 al N°{len(consolidado):03d}.")

    # Reordenar columnas: primero las del estándar semanal, luego extras
    cols_presentes  = [c for c in COLUMNAS_SEMANAL if c in consolidado.columns]
    cols_extra      = [c for c in consolidado.columns if c not in COLUMNAS_SEMANAL]
    consolidado     = consolidado[cols_presentes + cols_extra]

    log_fn(f"  ✓ Total consolidado: {len(consolidado)} filas.")
    return consolidado


def guardar_consolidado(df: pd.DataFrame, carpeta_salida: str, log_fn=print):
    """Guarda el consolidado como .xlsx (y opcionalmente .csv)."""
    try:
        fecha_ini = fecha_fin = "fecha"
        if "Fecha" in df.columns:
            fechas = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce").dropna()
            if not fechas.empty:
                fecha_ini = fechas.min().strftime("%d-%m-%Y")
                fecha_fin = fechas.max().strftime("%d-%m-%Y")

        nombre_base = f"Data_de_alertas_Semanal_{fecha_ini}_hasta_{fecha_fin}"
        ruta_xlsx   = os.path.join(carpeta_salida, nombre_base + ".xlsx")
        ruta_csv    = os.path.join(carpeta_salida, nombre_base + ".csv")

        with pd.ExcelWriter(ruta_xlsx, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=NOMBRE_HOJA_SALIDA, index=False)

        log_fn(f"\n  ✓ Excel guardado: {ruta_xlsx}")

        if GENERAR_CSV:
            df.to_csv(ruta_csv, index=False, encoding="utf-8-sig")
            log_fn(f"  ✓ CSV guardado:   {ruta_csv}")

        return ruta_xlsx

    except Exception as e:
        log_fn(f"  ✗ Error al guardar: {e}")
        return None


# ── Interfaz gráfica ───────────────────────────────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        root.title("Consolidador Semanal de Alertas")
        root.geometry("700x640")
        root.configure(bg="#0f1117")
        root.resizable(True, True)

        FONT_MONO = ("Courier New", 10)
        FONT_BOLD = ("Courier New", 11, "bold")
        BG     = "#0f1117"
        FG     = "#e2e8f0"
        ACCENT = "#818cf8"
        BTN_BG = "#4338ca"

        tk.Label(root, text="CONSOLIDADOR SEMANAL", font=("Courier New", 14, "bold"),
                 bg=BG, fg=ACCENT).pack(pady=(20, 2))
        tk.Label(root, text="Data de Alertas · Une archivos diarios en un reporte semanal",
                 font=FONT_MONO, bg=BG, fg="#64748b").pack(pady=(0, 16))

        # Botones selección
        sel_frame = tk.Frame(root, bg=BG)
        sel_frame.pack(pady=(0, 6))

        tk.Button(sel_frame, text="📂  Seleccionar archivos",
                  font=FONT_BOLD, bg=BTN_BG, fg="white",
                  activebackground="#3730a3", relief="flat", padx=12, pady=7,
                  command=self.seleccionar_archivos).pack(side="left", padx=6)

        tk.Button(sel_frame, text="📁  Seleccionar carpeta",
                  font=FONT_BOLD, bg="#0f766e", fg="white",
                  activebackground="#0d9488", relief="flat", padx=12, pady=7,
                  command=self.seleccionar_carpeta).pack(side="left", padx=6)

        self.lbl_seleccion = tk.Label(root, text="Ningún archivo seleccionado",
                                      font=FONT_MONO, bg=BG, fg="#475569")
        self.lbl_seleccion.pack(pady=(0, 10))

        # Lista de archivos seleccionados
        lista_frame = tk.Frame(root, bg=BG)
        lista_frame.pack(fill="x", padx=20, pady=(0, 8))

        self.lista_box = tk.Listbox(
            lista_frame, font=("Courier New", 9),
            bg="#161b27", fg="#94a3b8", selectbackground="#4338ca",
            relief="flat", bd=0, height=5
        )
        self.lista_box.pack(fill="x")

        # Carpeta de salida
        salida_frame = tk.Frame(root, bg=BG)
        salida_frame.pack(pady=(0, 10))

        tk.Label(salida_frame, text="Guardar en:", font=FONT_BOLD,
                 bg=BG, fg=FG).pack(side="left", padx=(0, 8))

        self.var_salida = tk.StringVar(
            value=os.path.expanduser("~") + os.sep + "Downloads"
        )
        tk.Entry(salida_frame, textvariable=self.var_salida, font=FONT_MONO,
                 width=36, bg="#1e2433", fg=FG, insertbackground=FG,
                 relief="flat", bd=4).pack(side="left")

        tk.Button(salida_frame, text="...", font=FONT_MONO,
                  bg="#334155", fg=FG, relief="flat", padx=6,
                  command=self.elegir_carpeta_salida).pack(side="left", padx=(4, 0))

        # Botón consolidar
        tk.Button(root, text="⚡  Consolidar y guardar",
                  font=("Courier New", 12, "bold"),
                  bg="#065f46", fg="white",
                  activebackground="#047857", relief="flat",
                  padx=20, pady=10,
                  command=self.consolidar).pack(pady=(4, 10))

        # Log
        tk.Label(root, text="LOG:", font=FONT_BOLD,
                 bg=BG, fg="#94a3b8").pack(anchor="w", padx=20)

        self.log = scrolledtext.ScrolledText(
            root, font=FONT_MONO, bg="#0d1117", fg="#86efac",
            insertbackground=FG, relief="flat", bd=0,
            padx=10, pady=10, wrap="word", height=12, state="disabled"
        )
        self.log.pack(fill="both", expand=True, padx=20, pady=(4, 20))

        self.rutas = []

    def log_write(self, texto):
        self.log.config(state="normal")
        self.log.insert("end", texto + "\n")
        self.log.see("end")
        self.log.config(state="disabled")
        self.root.update()

    def actualizar_lista(self):
        self.lista_box.delete(0, "end")
        for r in self.rutas:
            self.lista_box.insert("end", "  " + os.path.basename(r))
        self.lbl_seleccion.config(
            text=f"{len(self.rutas)} archivo(s) seleccionado(s)", fg="#818cf8"
        )

    def seleccionar_archivos(self):
        rutas = filedialog.askopenfilenames(
            title="Seleccionar archivos Excel diarios",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if rutas:
            self.rutas = sorted(list(rutas))
            self.actualizar_lista()

    def seleccionar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Carpeta con archivos diarios")
        if carpeta:
            self.rutas = sorted(
                glob.glob(os.path.join(carpeta, "*.xlsx")) +
                glob.glob(os.path.join(carpeta, "*.xls"))
            )
            self.actualizar_lista()

    def elegir_carpeta_salida(self):
        carpeta = filedialog.askdirectory(title="Carpeta donde guardar el consolidado")
        if carpeta:
            self.var_salida.set(carpeta)

    def consolidar(self):
        if not self.rutas:
            messagebox.showinfo("Info", "Primero selecciona los archivos o la carpeta.")
            return

        carpeta_salida = self.var_salida.get().strip()
        if not os.path.isdir(carpeta_salida):
            messagebox.showerror("Error", f"La carpeta de salida no existe:\n{carpeta_salida}")
            return

        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

        self.log_write(f"Iniciando consolidación de {len(self.rutas)} archivo(s)...\n")

        df = consolidar_archivos(self.rutas, log_fn=self.log_write)

        if df is None:
            messagebox.showerror("Error", "No se pudo generar el consolidado.")
            return

        ruta_salida = guardar_consolidado(df, carpeta_salida, log_fn=self.log_write)

        if ruta_salida:
            self.log_write("\n✅ ¡Proceso completado exitosamente!")
            messagebox.showinfo("Listo", f"Consolidado guardado en:\n{ruta_salida}")
        else:
            messagebox.showerror("Error", "No se pudo guardar el archivo.")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
