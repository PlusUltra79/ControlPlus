# ControlPlus CCO — App Android offline

> Gestión y seguimiento de unidades. 10 herramientas operativas en una APK **independiente, sin internet y sin navegador**.
> *Desarrollado por MoussaCorp · Se reservan todos los derechos.*

## Qué es
La versión Android de tu bot de Telegram. Misma lógica de negocio (`tools/*.py`, sin cambios), pero dentro del teléfono:
elige herramienta → adjunta el archivo del GTRMax → recibe mensajes listos para WhatsApp o un Excel.

## Arquitectura
```
APK
├─ Kotlin (MainActivity)   selector de archivos, guardar en Descargas, compartir, copiar
├─ Interfaz HTML/CSS/JS    dentro de la app (assets/www), animada y con ayudas por paso
├─ Chaquopy                Python 3.12 embebido
└─ Python
   ├─ api.py               reemplaza a bot.py: ejecutar(herramienta, params) -> JSON
   ├─ cp_compat.py         adaptadores para Android (tkinter, calamine, lxml)
   ├─ tools/               tus 10 herramientas originales
   └─ rutas_db.json
```
La app **no declara permiso INTERNET**. El WebView solo carga archivos locales y tiene las cargas de red bloqueadas.

## Compilar el APK gratis (GitHub Actions)
1. Crea un repositorio **privado** en GitHub y sube TODO el contenido de esta carpeta (incluida `.github/`).
2. Ve a la pestaña **Actions** → «Compilar APK» (se ejecuta solo al subir; o pulsa *Run workflow*).
3. Espera ~10-15 min (la primera vez descarga pandas para Android).
4. Entra a la ejecución terminada → **Artifacts** → `ControlPlusCCO-apk` → descarga el `.zip` → dentro está `ControlPlusCCO-v1.0.0.apk`.
5. Pasa el APK al teléfono, ábrelo y acepta «instalar apps de orígenes desconocidos».

El APK es *debug* (firmado con clave de depuración): sirve para instalar y usar. Para publicar en Play Store hay que firmarlo con tu propia clave (fase futura).

### Alternativa: Android Studio
Abre la carpeta, deja que sincronice y usa *Build → Build APK(s)*. Necesitas Python 3.12 instalado en la PC (Chaquopy lo usa al compilar).
Para probar en emulador: `gradle assembleDebug -Pabis=arm64-v8a,x86_64`.

## Cómo probar (checklist manual en el teléfono)
| # | Prueba | Resultado esperado |
|---|---|---|
| 1 | Abrir la app en modo avión | Splash, tutorial y «Motor listo» |
| 2 | Excesos: PDF de alarmas + hora | Un mensaje por unidad, botones Copiar/Compartir |
| 3 | Excesos: sin elegir PDF y pulsar Continuar | Mensaje «Elige el archivo para continuar» |
| 4 | Últimas UT: Reporte de Estatus (.xls) | Mensajes por grupo; probar «Obviar» con `R-066, 0343` |
| 5 | Ojo de Halcón: Estatus + Disponibilidad | Lista de incidencias |
| 6 | Data Alertas: «Cargar texto de ejemplo» → Generar | Excel; botón Guardar → aparece en Descargas/ControlPlus |
| 7 | Consolidador Semanal con 2+ Excel | Excel semanal + compartir |
| 8 | Desconexión: actualizar con tu Excel + PDF + fecha | Excel actualizado |
| 9 | Elegir un archivo equivocado | Mensaje de error entendible, sin cerrarse |
| 10 | Girar el teléfono / botón Atrás | Navega sin perder el estado |

## Pruebas automáticas (PC)
```
pip install -r tests/requirements-test.txt
CP_FORCE_FALLBACK=1 pytest -q tests/test_humo.py
CP_FORCE_FALLBACK=1 python tests/test_paridad.py <carpeta_con_tus_archivos> salida.json
```
`CP_FORCE_FALLBACK=1` simula Android (sin tkinter/calamine/lxml). Con tus archivos reales, el resultado en modo Android fue **idéntico** al modo normal en todas las herramientas probadas.

## Limitaciones conocidas
- **No se pudo compilar ni probar el APK en un teléfono real en este entorno.** La lógica Python sí se probó con tus archivos; el empaquetado (Chaquopy, instalación de pandas/pdfplumber en Android) se valida en tu primer build.
- Primera ejecución de cada sesión: 5-15 s mientras Python importa pandas.
- Solo arm64 (casi todos los teléfonos desde ~2017). Para x86/emulador ver arriba.
- Recorrido (herramienta 10) y Consolidador Mensual no se probaron con archivos reales (no se subieron muestras).
- APK de ~60-90 MB por incluir pandas/numpy.

## Mejoras futuras
Bloqueo por PIN, historial de reportes, edición de `rutas_db.json` desde la app, modo oscuro, firma de release para Play Store.
