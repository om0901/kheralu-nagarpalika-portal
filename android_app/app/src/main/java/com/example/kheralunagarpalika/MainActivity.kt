package com.example.kheralunagarpalika

import android.content.ContentValues
import android.content.Context
import android.content.Intent
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.os.Environment
import android.provider.MediaStore
import android.util.Base64
import android.webkit.JavascriptInterface
import android.webkit.WebChromeClient
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.Toast
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.ui.Modifier
import androidx.compose.ui.viewinterop.AndroidView
import java.io.File
import java.io.FileOutputStream

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContent {
            AndroidView(
                modifier = Modifier.fillMaxSize(),
                factory = { context ->
                    WebView(context).apply {
                        webViewClient = WebViewClient()
                        webChromeClient = WebChromeClient()
                        settings.javaScriptEnabled = true
                        settings.domStorageEnabled = true
                        settings.allowFileAccess = true
                        settings.allowContentAccess = true
                        settings.allowFileAccessFromFileURLs = true
                        settings.allowUniversalAccessFromFileURLs = true
                        settings.setSupportZoom(true)

                        addJavascriptInterface(WebAppInterface(context), "AndroidInterface")

                        loadUrl("file:///android_asset/index.html")
                    }
                }
            )
        }
    }
}

class WebAppInterface(private val mContext: Context) {
    @JavascriptInterface
    fun downloadFile(base64Data: String, fileName: String, mimeType: String) {
        try {
            val cleanBase64 = if (base64Data.contains(",")) base64Data.split(",")[1] else base64Data
            val fileBytes = Base64.decode(cleanBase64, Base64.DEFAULT)

            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                val resolver = mContext.contentResolver
                val contentValues = ContentValues().apply {
                    put(MediaStore.MediaColumns.DISPLAY_NAME, fileName)
                    put(MediaStore.MediaColumns.MIME_TYPE, mimeType)
                    put(MediaStore.MediaColumns.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS)
                }

                val uri = resolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, contentValues)
                uri?.let {
                    resolver.openOutputStream(it)?.use { outputStream ->
                        outputStream.write(fileBytes)
                    }
                    
                    Toast.makeText(mContext, "ફાઈલ ડાઉનલોડ થઈ ગઈ છે: $fileName", Toast.LENGTH_LONG).show()

                    val openIntent = Intent(Intent.ACTION_VIEW).apply {
                        setDataAndType(it, mimeType)
                        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                        addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)
                    }
                    mContext.startActivity(Intent.createChooser(openIntent, "ફાઈલ ખોલો"))
                }
            } else {
                val downloadsDir = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)
                val file = File(downloadsDir, fileName)
                FileOutputStream(file).use { it.write(fileBytes) }
                
                Toast.makeText(mContext, "ફાઈલ ડાઉનલોડ થઈ ગઈ છે: $fileName", Toast.LENGTH_LONG).show()

                val uri = Uri.fromFile(file)
                val openIntent = Intent(Intent.ACTION_VIEW).apply {
                    setDataAndType(uri, mimeType)
                    addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)
                }
                mContext.startActivity(Intent.createChooser(openIntent, "ફાઈલ ખોલો"))
            }
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(mContext, "Error: ${e.message}", Toast.LENGTH_SHORT).show()
        }
    }
}
