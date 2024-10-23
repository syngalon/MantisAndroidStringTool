package com.mantis.string.tools

import com.mantis.string.bean.XlsReadBean
import java.io.File
import java.io.FileOutputStream
import java.io.IOException

open class BaseSheet {
    init {
        val length = LANGUAGE_NAMES.size
        val languageMap: MutableMap<String, String> = HashMap()
        val folderMap: MutableMap<String, String> = HashMap()
        for (i in 0..<length) {
            languageMap[LANGUAGE_NAMES[i]] = LANGUAGE_FOLDERS[i]
            folderMap[LANGUAGE_FOLDERS[i]] = LANGUAGE_NAMES[i]
            LANGMAP = languageMap.entries
            FLOADER = folderMap.entries
        }
    }

    protected fun getFolderByLang(language: String?): String? {
        for ((key, value) in LANGMAP!!) {
            if (key.contains(language!!)) {
                return value
            }
        }
        return null
    }

    fun getLangByFolder(folderName: String): String? {
        for ((key, value) in FLOADER!!) {
            if (key == folderName) {
                return value
            }
        }
        return null
    }

    protected fun createFolder(rootPath: String?, path: String): String? {
        val paths = path.split("/".toRegex()).dropLastWhile { it.isEmpty() }.toTypedArray()
        val length = paths.size
        var currentPath = rootPath
        for (i in 0..<length) {
            val dir = paths[i]
            val file = File(currentPath, dir)
            if (!file.exists()) {
                file.mkdir()
            }
            currentPath = file.absolutePath
        }
        return currentPath
    }

    protected fun createFileAndData(path: String?, isArrayFile: Boolean, builder: XlsReadBean.Builder) {
        var file: File? = null
        file = if (isArrayFile) {
            File(path, builder.arrayName)
        } else {
            File(path, builder.stringName)
        }
        if (file.exists()) {
            file.delete()
        }
        var fos: FileOutputStream? = null
        try {
            file.createNewFile()
            fos = FileOutputStream(file)
            val sb = StringBuilder()
            sb.append(
                """<?xml version="1.0" encoding="utf-8"?>
<resources>"""
            ).append("\r\n")
            fos.write(sb.toString().toByteArray(charset("utf-8")))
        } catch (e: Exception) {
            // TODO Auto-generated catch block
            e.printStackTrace()
        } finally {
            if (fos != null) {
                try {
                    fos.close()
                } catch (e: IOException) {
                    // TODO Auto-generated catch block
                    e.printStackTrace()
                }
            }
        }
    }

    companion object {
        private val LANGUAGE_NAMES = arrayOf(
            "简体中文/簡中", "繁体中文/繁中", "English", "Czech",
            "Danish", "Dutch", "Spanish", "Finnish",
            "Portuguese", "French", "Deutsch", "Greek",
            "Italiano/Italian", "日语/Japanese", "Norwegian", "Polski/Polish",
            "Romanian", "Russian", "Swedish", "Turkish",
            "Arabic", "Chinese (Simple)", "Chinese (Traditional)", "Hungarian",
            "Thai", "Persian", "Vietnam/Vietnamese", "Korea/Korean",
            "Deutsch (German)"
        )
        private val LANGUAGE_FOLDERS = arrayOf(
            "values-zh-rCN", "values-zh-rTW", "values", "values-cs-rCZ",
            "values-da-rDK", "values-nl", "values-es", "values-fi-rFI",
            "values-pt", "values-fr", "values-de", "values-el-rGR",
            "values-it", "values-ja-rJP", "values-nb-rNO", "values-pl-rPL",
            "values-ro-rRO", "values-ru-rRU", "values-sv-rSE", "values-tr-rTR",
            "values-ar", "values-zh-rCN", "values-zh-rTW", "values-hu-rHU",
            "values-th-rTH", "values-fa", "values-vi-rVN", "values-ko-rKR",
            "values-de"
        )
        protected var LANGMAP: Set<Map.Entry<String, String>>? = null
        protected var FLOADER: Set<Map.Entry<String, String>>? = null
    }
}
