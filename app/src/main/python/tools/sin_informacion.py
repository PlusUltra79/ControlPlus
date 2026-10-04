"""
Sin Información - Generador de mensajes individuales para WhatsApp
------------------------------------------------------------------
Top 5 unidades con MAYOR distancia recorrida POR CADA SUB-EJE
(dentro de cada Eje). Ej: hasta 5 de Blv Acevedo, hasta 5 de Blv Brión,
hasta 5 de Serv Cont, hasta 5 de Serv Esp, etc.

Reglas de sub-eje:
  - Si Sub Flota es "Metro" o "Plaza Zamora" → se omite la línea Sub-eje en el mensaje.
  - Para todos los demás sub-ejes → se muestra "Sub-eje [valor]" en el mensaje.

Compatible con archivos .xls exportados por el sistema GTR.
Requiere: pip install beautifulsoup4 lxml

Uso: python sin_informacion.py
"""

from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
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

# Sub-flotas que NO se muestran en el mensaje (pero la unidad SÍ se incluye)
OCULTAR_SUBEJE = {"METRO", "PLAZA ZAMORA"}

# Normalización de nombres de sub-eje: la data GTR trae abreviaturas o
# variantes ("Ace", "Bri", "Tuy Ii", etc.) — aquí se corrigen a la
# escritura oficial que se debe mostrar en los mensajes.
NORMALIZAR_SUBEJE = {
    "NORTE":       "Norte",
    "CENTRO":      "Centro",
    "SUR":         "Sur",
    "ACEVEDO":     "Acevedo",
    "ACE":         "Acevedo",
    "BRION":       "Brion",
    "BRIÓN":       "Brion",
    "BRI":         "Brion",
    "BUROZ":       "Buroz",
    "BUR":         "Buroz",
    "PAEZ":        "Paez",
    "PÁEZ":        "Paez",
    "PA":          "Paez",
    "PEDRO GUAL":  "Pedro Gual",
    "PG":          "Pedro Gual",
    "TUY I":       "Tuy I",
    "TUY II":      "Tuy II",
    "SUCRE":       "Sucre",
    "ESP":         "Esp",
    "ESP.":        "Esp",
    "CONT":        "Cont",
    "CONT.":       "Cont",
}


def normalizar_subeje(subeje_up: str, subeje_raw: str) -> str:
    """
    Devuelve el nombre de sub-eje con la escritura oficial.
    Si el valor no está en el diccionario, cae de vuelta a un
    title-case simple (comportamiento anterior) como respaldo.
    """
    return NORMALIZAR_SUBEJE.get(subeje_up, subeje_raw.title())

TOP_N = 5  # Máximo de unidades por SUB-EJE (no por eje completo)

# ── Base de datos VIN ─────────────────────────────────────────────────────────
# Formato: "ALIAS": "VIN"  — agrega o edita según tu flota
VIN_DB = {
    "R-003": "4716", "R-022": "4919", "R-045": "3916", "R-046": "4625",
    "R-047": "9924", "R-048": "4791", "R-053": "8126", "R-060": "9867",
    "R-066": "2671", "R-069": "8396", "0219":  "8792", "0349":  "0589",
    "0489":  "8627", "0491":  "8644", "0566":  "6324", "0598":  "0415",
    "0613":  "0280", "0615":  "2295", "0647":  "0693", "0654":  "0722",
    "0664":  "0745", "R-010": "5944", "0352":  "0624", "0452":  "4436",
    "0599":  "0396", "0643":  "0705", "0645":  "0684", "0656":  "0723",
    "0665":  "0744", "0680":  "0179", "0721":  "0101", "0722":  "0052",
    "R-006": "9616", "R-007": "2273", "R-009": "7225", "R-052": "8252",
    "R-058": "2109", "0426":  "4361", "0522":  "0464", "0641":  "0703",
    "0657":  "0712", "0666":  "0713", "0700":  "0031", "R-005": "9299",
    "R-019": "5738", "0416":  "0132", "0465":  "4316", "0579":  "0426",
    "0616":  "2244", "R-049": "4622", "R-050": "4874", "0124":  "0947",
    "0343":  "0573", "0481":  "0781", "0482":  "4276", "0582":  "3270",
    "0609":  "6312", "0629":  "8060", "0636":  "0682", "0642":  "0691",
    "0646":  "0690", "0648":  "0692", "R-008": "8319", "R-034": "6159",
    "R-056": "0039", "0477":  "4434", "0667":  "0725", "R-012": "4649",
    "R-013": "4955", "R-014": "4742", "R-031": "4859", "0429":  "4326",
    "0495":  "8661", "0497":  "9519", "0500":  "9536", "0512":  "8436",
    "0516":  "7824", "0531":  "0404", "0546":  "6384", "0581":  "0423",
    "0597":  "0428", "0685":  "0257", "R-025": "4954", "R-033": "4665",
    "R-051": "3750", "0436":  "4458", "0511":  "9921", "0612":  "6362",
    "0632":  "7938", "R-039": "4778", "R-040": "0166", "R-041": "4765",
    "R-062": "1109", "0140":  "1182", "0424":  "4343", "0462":  "4412",
    "0463":  "4366", "0485":  "4413", "0536":  "0427", "0624":  "7851",
    "0625":  "7850", "0678":  "0242", "0705":  "0108", "R-042": "7819",
    "R-043": "4769", "R-044": "2905", "R-054": "3606", "R-055": "3949",
    "R-057": "7799", "R-063": "4704", "0417":  "4470", "0445":  "4273",
    "0450":  "4270", "0454":  "4505", "0614":  "2277", "R-001*":"4918",
    "R-016": "4728", "R-017": "2120", "R-018": "1419", "R-020": "8092",
    "R-021": "3679", "R-023": "4118", "R-024": "3419", "R-026": "4923",
    "R-027": "4718", "R-029": "4043", "R-061": "1481", "R-067": "8148",
    "0425":  "4352", "0448":  "4370", "0480":  "4414", "0484":  "4344",
    "0503":  "7815", "0583":  "0383", "0608":  "6346", "0637":  "0696",
    "0638":  "5153", "0446":  "4473", "0466":  "4489", "0592":  "2186",
    "0673":  "6627", "0690":  "0276", "0697":  "0275", "R-011": "4650",
    "0505":  "7852", "0514":  "9545", "0533":  "6307", "0538":  "0393",
    "0557":  "0385", "0562":  "6353", "0576":  "0454", "0591":  "0460",
    "0594":  "0452", "0671":  "6481", "0672":  "6490", "0683":  "0281",
    "0702":  "0016", "0712":  "0041", "R-064": "6656",
}

def get_vin(alias: str) -> str:
    """Busca el VIN del alias. Si no existe devuelve 'S/N'."""
    alias_limpio = alias.strip().upper()
    for k, v in VIN_DB.items():
        if k.upper() == alias_limpio:
            return v
    return "S/N"


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


# ── Lógica principal ──────────────────────────────────────────────────────────

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


def generar_mensajes(df):
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

    # Solo flotas válidas con distancia > 0 (sin excluir ninguna fila por sub-flota)
    df_f = df[df["_flota_up"].isin(set(ORDEN_FLOTAS)) & (df[col_dist] > 0)].copy()

    todos_mensajes = []

    import re as _re

    for flota_key in ORDEN_FLOTAS:
        nombre_eje  = NOMBRE_FLOTA.get(flota_key, flota_key)
        grupo_eje   = df_f[df_f["_flota_up"] == flota_key]

        # Sub-ejes presentes en este eje, en orden alfabético para salida
        # consistente (ej: dentro de Blv → Acevedo, Brión, Buroz, Páez...).
        subejes_presentes = sorted(grupo_eje["_subflota_up"].unique().tolist())

        mensajes_eje = []
        for subeje_up in subejes_presentes:
            # Top 5 (máximo) de ESTE sub-eje específico, no del eje completo
            grupo_sub = grupo_eje[grupo_eje["_subflota_up"] == subeje_up].nlargest(TOP_N, col_dist)

            for _, fila in grupo_sub.iterrows():
                subflota_raw = fila[col_subflota].strip()
                alias        = fila[col_alias].strip()
                distancia    = fila[col_dist]

                # Limpiar prefijo "UT:" que viene del archivo GTR
                alias = _re.sub(r"(?i)^ut:\s*", "", alias).strip()

                # Buscar VIN en la base de datos
                vin = get_vin(alias)

                # Nombre de sub-eje con escritura oficial (corrige abreviaturas)
                subeje_display = normalizar_subeje(subeje_up, subflota_raw)

                # Línea del eje: con o sin sub-eje según la regla
                if subeje_up in OCULTAR_SUBEJE:
                    linea_eje = f"• Eje *{nombre_eje}*"
                elif flota_key == "SERV.":
                    # Para Serv no se usa la palabra "Sub-eje": se muestra
                    # directamente como "Eje *Serv* *Cont:*" / "*Esp:*"
                    linea_eje = f"• Eje *{nombre_eje}* *{subeje_display}:*"
                else:
                    linea_eje = f"• Eje *{nombre_eje}* Sub-eje *{subeje_display}*"

                msg = (
                    f"Buenos Días\n"
                    f"{linea_eje}\n"
                    f"• Unidad *{alias}* - Vin *{vin}*\n"
                    f"• CTS: *S.I*\n"
                    f"\n"
                    f"*Incidencia:* Sin información\n"
                    f"*Observación:* No se ha recibido reporte de disponibilidad "
                    f"*({distancia:.2f}km recorridos).*"
                )
                mensajes_eje.append(msg)

        todos_mensajes.append((nombre_eje, mensajes_eje))

    return todos_mensajes


def formatear_salida(todos_mensajes):
    bloques = []
    for nombre_eje, mensajes in todos_mensajes:
        if not mensajes:
            bloques.append(
                f"{'═'*50}\n  EJE {nombre_eje.upper()} — Sin unidades\n{'═'*50}"
            )
            continue
        encabezado = f"{'═'*50}\n  EJE {nombre_eje.upper()} — {len(mensajes)} unidades (Top 5 por sub-eje)\n{'═'*50}"
        cuerpo = ("\n" + "─"*50 + "\n").join(mensajes)
        bloques.append(encabezado + "\n" + cuerpo)
    return "\n\n".join(bloques)


# ── Interfaz gráfica ──────────────────────────────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        root.title("Sin Información - Mensajes WhatsApp")
        root.geometry("660x640")
        root.configure(bg="#0f1117")
        root.resizable(True, True)

        FONT_MONO = ("Courier New", 10)
        FONT_BOLD = ("Courier New", 11, "bold")
        BG     = "#0f1117"
        FG     = "#e2e8f0"
        ACCENT = "#f59e0b"
        BTN_BG = "#b45309"

        tk.Label(root, text="SIN INFORMACIÓN", font=("Courier New", 14, "bold"),
                 bg=BG, fg=ACCENT).pack(pady=(20, 2))
        tk.Label(root, text="Top 5 por sub-eje · Mensajes individuales para WhatsApp",
                 font=FONT_MONO, bg=BG, fg="#64748b").pack(pady=(0, 16))

        tk.Button(root, text="📂  Seleccionar archivo Excel (.xls / .xlsx)",
                  font=FONT_BOLD, bg=BTN_BG, fg="white",
                  activebackground="#92400e", activeforeground="white",
                  relief="flat", padx=16, pady=8,
                  command=self.cargar_archivo).pack(pady=(0, 6))

        self.lbl_archivo = tk.Label(root, text="Ningún archivo seleccionado",
                                    font=FONT_MONO, bg=BG, fg="#475569")
        self.lbl_archivo.pack(pady=(0, 12))

        tk.Label(root, text="MENSAJES GENERADOS:", font=FONT_BOLD,
                 bg=BG, fg="#94a3b8").pack(anchor="w", padx=20)

        self.txt = scrolledtext.ScrolledText(
            root, font=FONT_MONO, bg="#0d1117", fg="#fde68a",
            insertbackground=FG, relief="flat", bd=0,
            padx=10, pady=10, wrap="word", height=22
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

        self.todos_mensajes = []

    def cargar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo",
            filetypes=[("Archivos Excel/GTR", "*.xls *.xlsx")]
        )
        if not ruta:
            return

        self.lbl_archivo.config(text=os.path.basename(ruta), fg="#f59e0b")

        try:
            df = leer_archivo(ruta)
            self.todos_mensajes = generar_mensajes(df)
            texto = formatear_salida(self.todos_mensajes)
            self.txt.delete("1.0", "end")
            self.txt.insert("end", texto)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def copiar_todo(self):
        if not self.todos_mensajes:
            messagebox.showinfo("Info", "No hay mensajes generados.")
            return
        todos = []
        for _, mensajes in self.todos_mensajes:
            todos.extend(mensajes)
        self.root.clipboard_clear()
        self.root.clipboard_append("\n\n".join(todos))
        messagebox.showinfo("Listo", f"¡{len(todos)} mensajes copiados al portapapeles!")

    def guardar_txt(self):
        if not self.todos_mensajes:
            messagebox.showinfo("Info", "No hay mensajes generados.")
            return
        ruta = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Archivo de texto", "*.txt")],
            initialfile=f"sin_informacion_{datetime.now().strftime('%d%m%Y_%H%M')}.txt"
        )
        if not ruta:
            return
        todos = []
        for nombre_eje, mensajes in self.todos_mensajes:
            if mensajes:
                todos.append(f"=== EJE {nombre_eje.upper()} ===")
                todos.extend(mensajes)
        with open(ruta, "w", encoding="utf-8") as f:
            f.write("\n\n".join(todos))
        messagebox.showinfo("Guardado", f"Archivo guardado en:\n{ruta}")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
