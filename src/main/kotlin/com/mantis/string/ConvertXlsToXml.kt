package com.mantis.string

import com.mantis.string.bean.XlsReadBean
import com.mantis.string.tools.ReadXlsManager
import java.io.File


object ConvertXlsToXml {
    private const val ROOT_FILE = "sample.xls"
    private const val DIR_NAME = "Sample_values"

    /**
     * file to save strings
     */
    private const val STRING_NAME = "test_strings.xml"

    /**
     * file to save arrays
     */
    private const val ARRAY_NAME = "test_arrays.xml"
    private const val IGNORE_ROW = 1
    private var ROOT_PATH: String? = null
    @JvmStatic
    fun main(args: Array<String>) {
        val file = File("")
        ROOT_PATH = file.absolutePath
        val bean: XlsReadBean = XlsReadBean.Builder()
            .setRootPath(ROOT_PATH)
            .setXlsFile(ROOT_FILE)
            .setFileFolderName(DIR_NAME)
            .setIgnoreRow(IGNORE_ROW)
            .setStringName(STRING_NAME)
            .setArrayName(ARRAY_NAME)
            .builder()
        ReadXlsManager.Holder.instance.readXls(bean.builder)
        println("file converted: " + ROOT_PATH + File.separator + DIR_NAME)
    }
}
