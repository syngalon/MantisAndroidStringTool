package com.mantis.string

import com.mantis.string.bean.XlsWriteBean
import com.mantis.string.tools.WriteXlsManager
import java.io.File


object CovertXmlToXls {
    /**
     * the direction where value existed
     */
    private const val VALUE_PATH = "values_to_convert"

    /**
     * file to be converted
     */
    private const val XLS_NAME = "convertedFile.xls"

    /**
     * current direction
     */
    private var ROOT_PATH: String? = null
    @JvmStatic
    fun main(args: Array<String>) {
        val file = File("")
        ROOT_PATH = file.absolutePath
        val bean: XlsWriteBean = XlsWriteBean.Builder()
            .setRootPath(ROOT_PATH)
            .setFileFolderName(VALUE_PATH)
            .setXlsName(XLS_NAME)
            .builder()

        WriteXlsManager.Holder.instance.startWrite(bean.builder)
        println("file converted: " + ROOT_PATH + File.separator + VALUE_PATH)
    }
}
