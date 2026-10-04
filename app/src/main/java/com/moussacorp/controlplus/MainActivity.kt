package com.moussacorp.controlplus

import android.annotation.SuppressLint
import android.content.ClipData
import android.content.ClipboardManager
import android.content.ContentValues
import android.content.Context
import android.content.Intent
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.os.Environment
import android.provider.MediaStore
import android.provider.OpenableColumns
import android.webkit.JavascriptInterface
import android.webkit.WebSettings
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.Toast
import androidx.activity.ComponentActivity
import androidx.activity.OnBackPressedCallback
import androidx.activity.result.contract.ActivityResultContracts
import androidx.core.content.FileProvider
import com.chaquo.python.Python
import org.json.JSONArray
import org.json.JSONObject
import java.io.File
import java.util.concurrent.Executors

/**
 * ControlPlus CCO — cascarón Android.
 * Interfaz (HTML/CSS/JS) embebida en un WebView local + Python embebido (Chaquopy).
 * No declara permiso INTERNET y bloquea cargas de red: 100 % offline.
 */
class MainActivity : ComponentActivity() {

    private lateinit var webView: WebView
    private val ejecutor = Executors.newSingleThreadExecutor()
    private var idSelector: Int = -1

    private val selectorUnico =
        registerForActivityResult(ActivityResultContracts.OpenDocument()) { uri ->
            entregarArchivos(if (uri != null) listOf(uri) else emptyList())
        }
    private val selectorMultiple =
        registerForActivityResult(ActivityResultContracts.OpenMultipleDocuments()) { uris ->
            entregarArchivos(uris)
        }

    @SuppressLint("SetJavaScriptEnabled")
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        limpiarTemporales()

        webView = WebView(this).apply {
            settings.javaScriptEnabled = true
            settings.domStorageEnabled = true
            settings.allowFileAccess = true          // solo para file:///android_asset
            settings.allowContentAccess = false
            settings.blockNetworkLoads = true        // sin red, por diseño
            settings.cacheMode = WebSettings.LOAD_NO_CACHE
            settings.mediaPlaybackRequiresUserGesture = true
            webViewClient = WebViewClient()
            addJavascriptInterface(Puente(), "Android")
        }
        setContentView(webView)
        webView.loadUrl("file:///android_asset/www/index.html")

        onBackPressedDispatcher.addCallback(this, object : OnBackPressedCallback(true) {
            override fun handleOnBackPressed() {
                webView.evaluateJavascript("(window.__atras ? window.__atras() : false)") { r ->
                    if (r != "true") {
                        isEnabled = false
                        onBackPressedDispatcher.onBackPressed()
                    }
                }
            }
        })

        // Precalienta Python (importar pandas tarda unos segundos la primera vez).
        ejecutor.submit {
            val estado = try {
                Python.getInstance().getModule("api").callAttr("estado").toString()
            } catch (e: Throwable) {
                JSONObject().put("error", e.message ?: e.toString()).toString()
            }
            enviarAJs(0, estado)
        }
    }

    override fun onDestroy() {
        super.onDestroy()
        ejecutor.shutdown()
    }

    // ── Utilidades ───────────────────────────────────────────────────────────
    private fun enviarAJs(id: Int, json: String) {
        webView.post {
            webView.evaluateJavascript("window.__nativo($id, ${JSONObject.quote(json)})", null)
        }
    }

    private fun limpiarTemporales() {
        File(cacheDir, "in").deleteRecursively()
        File(cacheDir, "out").deleteRecursively()
    }

    private fun nombreDe(uri: Uri): String {
        contentResolver.query(uri, null, null, null, null)?.use { c ->
            val i = c.getColumnIndex(OpenableColumns.DISPLAY_NAME)
            if (i >= 0 && c.moveToFirst()) return c.getString(i)
        }
        return uri.lastPathSegment?.substringAfterLast('/') ?: "archivo"
    }

    private fun entregarArchivos(uris: List<Uri>) {
        val id = idSelector
        ejecutor.submit {
            val lista = JSONArray()
            try {
                uris.forEachIndexed { n, uri ->
                    val nombre = nombreDe(uri).replace(Regex("[\\\\/:*?\"<>|]"), "_")
                    val carpeta = File(cacheDir, "in/${System.currentTimeMillis()}_$n").apply { mkdirs() }
                    val destino = File(carpeta, nombre)
                    contentResolver.openInputStream(uri)?.use { entrada ->
                        destino.outputStream().use { salida -> entrada.copyTo(salida) }
                    }
                    lista.put(JSONObject().put("nombre", nombre).put("ruta", destino.absolutePath).put("bytes", destino.length()))
                }
                enviarAJs(id, JSONObject().put("ok", true).put("archivos", lista).toString())
            } catch (e: Exception) {
                enviarAJs(id, JSONObject().put("ok", false).put("error", "No se pudo abrir el archivo: ${e.message}").toString())
            }
        }
    }

    private fun mimeDe(nombre: String): String = when (nombre.substringAfterLast('.').lowercase()) {
        "xlsx" -> "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        "csv" -> "text/csv"
        "pdf" -> "application/pdf"
        else -> "application/octet-stream"
    }

    /** Solo se permite exportar/compartir archivos generados por la propia app. */
    private fun rutaPermitida(ruta: String): File? {
        val f = File(ruta).canonicalFile
        val base = File(cacheDir, "out").canonicalFile
        return if (f.exists() && f.path.startsWith(base.path + File.separator)) f else null
    }

    // ── Puente JS → Kotlin ───────────────────────────────────────────────────
    inner class Puente {

        @JavascriptInterface
        fun elegirArchivos(id: Int, multiple: Boolean) {
            idSelector = id
            runOnUiThread {
                val tipos = arrayOf("*/*")
                if (multiple) selectorMultiple.launch(tipos) else selectorUnico.launch(tipos)
            }
        }

        @JavascriptInterface
        fun ejecutar(id: Int, herramienta: String, paramsJson: String) {
            ejecutor.submit {
                val respuesta = try {
                    val params = JSONObject(paramsJson)
                    val salida = File(cacheDir, "out/${System.currentTimeMillis()}").apply { mkdirs() }
                    params.put("salida", salida.absolutePath)
                    Python.getInstance().getModule("api")
                        .callAttr("ejecutar", herramienta, params.toString()).toString()
                } catch (e: Throwable) {
                    JSONObject().put("ok", false)
                        .put("error", "Error interno: ${e.message ?: e.toString()}").toString()
                }
                enviarAJs(id, respuesta)
            }
        }

        @JavascriptInterface
        fun unidadConocida(id: Int, unidad: String) {
            ejecutor.submit {
                val r = try {
                    Python.getInstance().getModule("api").callAttr("unidad_conocida", unidad).toString()
                } catch (e: Throwable) {
                    "{\"alias\":null}"
                }
                enviarAJs(id, r)
            }
        }

        @JavascriptInterface
        fun copiar(texto: String) {
            runOnUiThread {
                val cm = getSystemService(Context.CLIPBOARD_SERVICE) as ClipboardManager
                cm.setPrimaryClip(ClipData.newPlainText("ControlPlus", texto))
            }
        }

        @JavascriptInterface
        fun compartirTexto(texto: String) {
            runOnUiThread {
                val i = Intent(Intent.ACTION_SEND).apply {
                    type = "text/plain"
                    putExtra(Intent.EXTRA_TEXT, texto)
                }
                startActivity(Intent.createChooser(i, "Compartir reporte"))
            }
        }

        @JavascriptInterface
        fun compartirArchivo(ruta: String) {
            val f = rutaPermitida(ruta) ?: return
            runOnUiThread {
                val uri = FileProvider.getUriForFile(this@MainActivity, "$packageName.files", f)
                val i = Intent(Intent.ACTION_SEND).apply {
                    type = mimeDe(f.name)
                    putExtra(Intent.EXTRA_STREAM, uri)
                    addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                }
                startActivity(Intent.createChooser(i, "Compartir archivo"))
            }
        }

        /** Guarda en Descargas/ControlPlus. Devuelve JSON {ok, nombre, carpeta | error}. */
        @JavascriptInterface
        fun guardarArchivo(ruta: String): String {
            val f = rutaPermitida(ruta)
                ?: return JSONObject().put("ok", false).put("error", "Archivo no disponible").toString()
            return try {
                if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                    val valores = ContentValues().apply {
                        put(MediaStore.Downloads.DISPLAY_NAME, f.name)
                        put(MediaStore.Downloads.MIME_TYPE, mimeDe(f.name))
                        put(MediaStore.Downloads.RELATIVE_PATH, "${Environment.DIRECTORY_DOWNLOADS}/ControlPlus")
                    }
                    val uri = contentResolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, valores)
                        ?: throw IllegalStateException("No se pudo crear el archivo en Descargas")
                    contentResolver.openOutputStream(uri)?.use { o -> f.inputStream().use { it.copyTo(o) } }
                } else {
                    val dir = File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), "ControlPlus")
                    dir.mkdirs()
                    f.copyTo(File(dir, f.name), overwrite = true)
                }
                JSONObject().put("ok", true).put("nombre", f.name).put("carpeta", "Descargas/ControlPlus").toString()
            } catch (e: Exception) {
                JSONObject().put("ok", false).put("error", e.message ?: "No se pudo guardar").toString()
            }
        }

        @JavascriptInterface
        fun aviso(texto: String) {
            runOnUiThread { Toast.makeText(this@MainActivity, texto, Toast.LENGTH_SHORT).show() }
        }

        @JavascriptInterface
        fun salir() {
            runOnUiThread { finish() }
        }
    }
}
