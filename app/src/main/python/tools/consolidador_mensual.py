"""
Consolidador Mensual de Reportes
----------------------------------
Lee varios archivos Excel SEMANALES y genera un único archivo consolidado mensual.

Cada archivo semanal debe tener una hoja con los datos consolidados.
El script busca automáticamente la hoja correcta (la que contenga las columnas
esperadas), sin importar cómo se llame ("General de Marzo", "General", etc.)

Columnas esperadas en cada semanal:
  Responsable, Reporte, Fecha, Hora, Eje, Sub Eje, Unidad, Cts,
  Tipo de reporte, Incidencias, Niveles, Descripción, Observación,
  Respuesta, Detalle

Requiere: pip install pandas openpyxl
          tkinter viene incluido con Python en Windows

Uso: python consolidador_mensual.py
"""

import os, re, glob
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# ── Configuración ──────────────────────────────────────────────────────────────
# Columnas que debe tener la hoja de datos del semanal
COLUMNAS_ESPERADAS = {"Responsable", "Reporte", "Fecha", "Eje", "Unidad"}

# Orden exacto de columnas en el archivo mensual de salida
COLUMNAS_MENSUAL = [
    "Responsable", "Reporte", "Fecha", "Hora",
    "Eje", "Sub Eje", "Unidad", "Cts",
    "Tipo de reporte", "Incidencias", "Niveles",
    "Descripción", "Observación", "Respuesta", "Detalle"
]

NOMBRE_HOJA_SALIDA = "General"
GENERAR_CSV        = True

# ── Funciones ──────────────────────────────────────────────────────────────────

def encontrar_hoja_datos(ruta_archivo: str) -> str | None:
    """
    Busca automáticamente la hoja que contiene las columnas de datos.
    Retorna el nombre de la hoja o None si no encuentra ninguna.
    """
    xl = pd.ExcelFile(ruta_archivo, engine="openpyxl")
    for hoja in xl.sheet_names:
        try:
            df_prueba = pd.read_excel(xl, sheet_name=hoja, nrows=2)
            if COLUMNAS_ESPERADAS.issubset(set(df_prueba.columns)):
                return hoja
        except Exception:
            continue
    return None


def leer_archivo_semanal(ruta_archivo: str, log_fn=print):
    """
    Lee la hoja de datos de un archivo semanal.
    Retorna un DataFrame limpio o None si hay error.
    """
    nombre = os.path.basename(ruta_archivo)
    try:
        hoja = encontrar_hoja_datos(ruta_archivo)
        if not hoja:
            log_fn(f"  ✗ {nombre}: no se encontró hoja con columnas de datos. Se omite.")
            return None

        df = pd.read_excel(ruta_archivo, sheet_name=hoja, engine="openpyxl")

        if df.empty:
            log_fn(f"  ⚠ {nombre}: hoja '{hoja}' vacía, se omite.")
            return None

        # Eliminar la columna Reporte original (se renumerará globalmente)
        if "Reporte" in df.columns:
            df.drop(columns=["Reporte"], inplace=True)

        log_fn(f"  ✓ {nombre} [hoja: '{hoja}']: {len(df)} filas leídas.")
        return df

    except Exception as e:
        log_fn(f"  ✗ {nombre}: ERROR → {e}")
        return None


def consolidar_semanales(lista_rutas: list, log_fn=print):
    """
    Lee todos los semanales, los concatena, limpia, ordena y renumera.
    """
    dataframes = []

    for ruta in lista_rutas:
        df = leer_archivo_semanal(ruta, log_fn)
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

    # Ordenar por Fecha (de más antigua a más reciente),
    # respetando el orden original dentro de cada fecha
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

    # Renumerar Reporte de forma consecutiva global
    consolidado.insert(1, "Reporte", [f"N°{i:03d}" for i in range(1, len(consolidado) + 1)])
    log_fn(f"  ℹ Columna 'Reporte' numerada del N°001 al N°{len(consolidado):03d}.")

    # Reordenar columnas al estándar mensual
    cols_presentes = [c for c in COLUMNAS_MENSUAL if c in consolidado.columns]
    cols_extra     = [c for c in consolidado.columns if c not in COLUMNAS_MENSUAL]
    consolidado    = consolidado[cols_presentes + cols_extra]

    log_fn(f"  ✓ Total consolidado: {len(consolidado)} filas.")
    return consolidado


def guardar_consolidado(df: pd.DataFrame, carpeta_salida: str, log_fn=print):
    """Guarda el consolidado mensual como .xlsx (y opcionalmente .csv)."""
    try:
        fecha_ini = fecha_fin = "fecha"
        if "Fecha" in df.columns:
            fechas = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce").dropna()
            if not fechas.empty:
                fecha_ini = fechas.min().strftime("%d-%m-%Y")
                fecha_fin = fechas.max().strftime("%d-%m-%Y")

        nombre_base = f"Data_de_alertas_Mensual_{fecha_ini}_hasta_{fecha_fin}"
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
        root.title("Consolidador Mensual de Alertas")
        root.geometry("700x640")
        root.configure(bg="#0f1117")
        root.resizable(True, True)

        FONT_MONO = ("Courier New", 10)
        FONT_BOLD = ("Courier New", 11, "bold")
        BG     = "#0f1117"
        FG     = "#e2e8f0"
        ACCENT = "#f472b6"
        BTN_BG = "#9d174d"

        tk.Label(root, text="CONSOLIDADOR MENSUAL", font=("Courier New", 14, "bold"),
                 bg=BG, fg=ACCENT).pack(pady=(20, 2))
        tk.Label(root, text="Data de Alertas · Une archivos semanales en un reporte mensual",
                 font=FONT_MONO, bg=BG, fg="#64748b").pack(pady=(0, 16))

        # Botones selección
        sel_frame = tk.Frame(root, bg=BG)
        sel_frame.pack(pady=(0, 6))

        tk.Button(sel_frame, text="📂  Seleccionar archivos semanales",
                  font=FONT_BOLD, bg=BTN_BG, fg="white",
                  activebackground="#831843", relief="flat", padx=12, pady=7,
                  command=self.seleccionar_archivos).pack(side="left", padx=6)

        tk.Button(sel_frame, text="📁  Seleccionar carpeta",
                  font=FONT_BOLD, bg="#0f766e", fg="white",
                  activebackground="#0d9488", relief="flat", padx=12, pady=7,
                  command=self.seleccionar_carpeta).pack(side="left", padx=6)

        self.lbl_seleccion = tk.Label(root, text="Ningún archivo seleccionado",
                                      font=FONT_MONO, bg=BG, fg="#475569")
        self.lbl_seleccion.pack(pady=(0, 6))

        # Lista de archivos seleccionados
        self.lista_box = tk.Listbox(
            root, font=("Courier New", 9),
            bg="#161b27", fg="#94a3b8", selectbackground="#9d174d",
            relief="flat", bd=0, height=5
        )
        self.lista_box.pack(fill="x", padx=20, pady=(0, 10))

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
            padx=10, pady=10, wrap="word", height=13, state="disabled"
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
            text=f"{len(self.rutas)} archivo(s) seleccionado(s)", fg="#f472b6"
        )

    def seleccionar_archivos(self):
        rutas = filedialog.askopenfilenames(
            title="Seleccionar archivos Excel semanales",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if rutas:
            self.rutas = sorted(list(rutas))
            self.actualizar_lista()

    def seleccionar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Carpeta con archivos semanales")
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

        self.log_write(f"Iniciando consolidación mensual de {len(self.rutas)} semanal(es)...\n")

        df = consolidar_semanales(self.rutas, log_fn=self.log_write)

        if df is None:
            messagebox.showerror("Error", "No se pudo generar el consolidado.")
            return

        ruta_salida = guardar_consolidado(df, carpeta_salida, log_fn=self.log_write)

        if ruta_salida:
            self.log_write("\n✅ ¡Proceso completado exitosamente!")
            messagebox.showinfo("Listo", f"Consolidado mensual guardado en:\n{ruta_salida}")
        else:
            messagebox.showerror("Error", "No se pudo guardar el archivo.")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
