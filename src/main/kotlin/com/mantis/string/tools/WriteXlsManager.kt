package com.mantis.string.tools

import com.mantis.string.bean.Folder
import com.mantis.string.bean.StringRow
import com.mantis.string.bean.XlsWriteBean
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.dom4j.Document
import org.dom4j.Element
import org.dom4j.io.SAXReader
import java.io.File
import java.io.FileOutputStream
import java.io.IOException
import java.util.*
import kotlin.collections.ArrayList


class WriteXlsManager private constructor() : BaseSheet() {
    private var isArrayXls = false
    private val keyList: MutableList<String> = ArrayList()
    private var workbook: Workbook? = null

    object Holder {
        var instance = WriteXlsManager()
    }

    fun startWrite(builder: XlsWriteBean.Builder) {
        println("fileFolderName1: " + builder.fileFolderName)
        val file: File? = builder.fileFolderName?.let { File(builder.rootPath, it) }
        var fos: FileOutputStream? = null
        if (file?.exists() == true) {
            try {
                println("xlsName: ${builder.xlsName}")
                workbook = if (builder.xlsName.lowercase(Locale.getDefault()).endsWith("xlsx")) {
                    XSSFWorkbook()
                } else {
                    HSSFWorkbook()
                }
                val filePath: String = builder.rootPath + File.separator + builder.fileFolderName
                val xlsFile = File(filePath, builder.xlsName)
                if (xlsFile.exists()) {
                    xlsFile.delete()
                }
                startWriteWorkbook(file)
                fos = FileOutputStream(File(filePath, builder.xlsName))
                workbook?.write(fos)
                fos.close()
            } catch (e: Exception) {
                // TODO: handle exception
                println("error: $e")
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
        } else {
            println(("cannot find " + builder.rootPath) + File.separator + " " + builder.fileFolderName)
        }
    }

    /**
     * begin to write data
     * @param rootFile
     */
    private fun startWriteWorkbook(rootFile: File) {
        if (rootFile.exists()) {
            val cellStyle: CellStyle? = workbook?.createCellStyle()
            cellStyle?.wrapText = true
            val createHelper: CreationHelper? = workbook?.creationHelper
            val thisFolder: Folder = getFolder(rootFile)
            workbook?.createSheet()
            if (thisFolder.folderPaths?.isNotEmpty() == true) {
                for (folderPath in thisFolder.folderPaths!!) {
                    val folder = File(folderPath)
                    if (folder.exists()) {
                        val valueNames = getFolderFileNameList(folderPath)
                        for (valueName in valueNames) {
                            println("valueName: $valueName")
                            var sheet: Sheet?
                            var row: Row?
                            var startColumn = 1
                            val size = valueNames.size
                            if (valueName.contains("array")) {
                                isArrayXls = true
                                if (size == 1) {
                                    // if only array, begins at 0
                                    workbook?.setSheetName(0, ARRAY_NAME)
                                    sheet = workbook?.getSheet(ARRAY_NAME)
                                } else {
                                    sheet = workbook?.getSheet(ARRAY_NAME)
                                    if (sheet == null) {
                                        sheet = workbook?.createSheet(ARRAY_NAME)
                                    }
                                }
                                row = sheet?.createRow(0)
                                row?.createCell(0)?.setCellValue(ARRAY_TYPE_DECLARE)
                            } else {
                                isArrayXls = false
                                sheet = workbook?.getSheet(STRING_NAME)
                                if (sheet == null) {
                                    workbook?.setSheetName(0, STRING_NAME)
                                    sheet = workbook?.getSheet(STRING_NAME)
                                }
                                row = sheet?.createRow(0)
                                row?.createCell(0)?.setCellValue(STRING_TYPE_DECLARE)
                            }
                            sheet?.setColumnWidth(0, 30 * 256)
                            if (isArrayXls) {
                                startColumn = 2
                            }

                            for (langIndex in 0..<thisFolder.languages?.size!!) {
                                row?.createCell(langIndex + startColumn)?.setCellValue(
                                    createHelper?.createRichTextString(thisFolder.languages!![langIndex])
                                )
                                val lists: List<StringRow> = createHelper?.let {
                                    parseStringXml(
                                        thisFolder.folderPaths!![langIndex],
                                        valueName, it
                                    )
                                }!!
                                if (lists.isNotEmpty()) {
                                    writeDataToXls(lists, sheet, langIndex)
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    private fun getFolder(rootFile: File): Folder {
        val folder = Folder()
        folder.path = rootFile.absolutePath
        folder.languages = ArrayList()
        folder.folderPaths = ArrayList()
        if (rootFile.exists()) {
            val files = rootFile.listFiles()
            for (file in files!!) {
                getLangByFolder(getFolderName(file)!!)?.let { (folder.languages as ArrayList<String>).add(it) }
                (folder.folderPaths as ArrayList<String>).add(file.absolutePath)
            }
        }
        return folder
    }

    /**
     * write data to xls
     */
    private fun writeDataToXls(lists: List<StringRow>, sheet: Sheet?, rIndex: Int) {
        val itemCountList: MutableList<Int> = ArrayList()
        val cellStyle: CellStyle? = workbook?.createCellStyle()
        cellStyle?.wrapText = true
        val size = lists.size

        for (cIndex in 0..<size) {
            val stringRow: StringRow = lists[cIndex]

            if (!isArrayXls) {
                if (rIndex == 0) {
                    val valueRow: Row? = sheet?.createRow(cIndex + 1)
                    valueRow?.createCell(0)?.setCellValue(stringRow.key)
                    val cell: Cell? = valueRow?.createCell(rIndex + 1)
                    cell?.cellStyle = cellStyle
                    sheet?.setColumnWidth(cIndex + 1, 25 * 256)
                    cell?.setCellValue(stringRow.value)
                    stringRow.key?.let { keyList.add(it) }
                } else {
                    val index = stringRow.key?.let { getIndexFromKey(it) }
                    if (index != -1) {

                        index?.let {
                            var valueRow: Row? = sheet?.getRow(index + 1)
                            if (valueRow == null) {
                                valueRow = sheet?.createRow(index + 1)
                            }
                            val cell: Cell? = valueRow?.createCell(rIndex + 1)
                            cell?.cellStyle = cellStyle
                            sheet?.setColumnWidth(index, 35 * 256)
                            cell?.setCellValue(stringRow.value)
                        }
                    }
                }
            } else {
                var index = cIndex + 1
                if (cIndex > 0) {
                    var num = 0
                    for (itemIndex in itemCountList.indices) {
                        num += itemCountList[itemIndex]
                    }
                    index = cIndex + 1 + num
                }
                val valueRow: Row? = sheet?.createRow(index)
                if (stringRow.key != null) {
                    valueRow?.createCell(0)?.setCellValue(stringRow.key)
                    if (stringRow.items != null) {
                        val count: Int = stringRow.items!!.size - 1
                        itemCountList.add(count)
                        index += 1
                        println("count: $count ")
                        for (cRow in 0..<count) {
                            val item: String = stringRow.items!![cRow]
                            val itemIndex = index + cRow
                            var itemRow: Row?
                            sheet?.let {
                                itemRow = sheet.getRow(itemIndex)
                                if (itemRow == null) {
                                    itemRow = sheet.createRow(itemIndex)
                                }
                                itemRow?.createCell(1)?.setCellValue("<item>")
                                val cell: Cell? = itemRow?.createCell(rIndex + 2)
                                cell?.cellStyle = cellStyle
                                sheet.setColumnWidth(cIndex + 1, 35 * 256)
                                cell?.setCellValue(item)
                                println(("item: " + itemRow?.getCell(0)) + " " + itemRow?.getCell(1))
                            }
                        }
                    }
                }
            }
        }
    }

    private fun parseStringXml(path: String, stringName: String, createHelper: CreationHelper): List<StringRow> {
        val lists: MutableList<StringRow> = ArrayList()
        try {
            val file = File(path, stringName)
            if (file.exists()) {
                val read = SAXReader()
                val document: Document = read.read(file)
                val root: Element = document.rootElement

                val it: Iterator<Element> = root.elementIterator()
                while (it.hasNext()) {
                    val element: Element = it.next()

                    val cusRow = StringRow()
                    cusRow.key = element.attributeValue("name")
                    cusRow.value = element.stringValue
                    if (isArrayXls) {
                        if (cusRow.value != null) {
                            val items: List<String> = cusRow.value!!.split("\n")
                            cusRow.items = ArrayList()
                            for (item in items) {
                                if (item.length > 1) {
                                    (cusRow.items as ArrayList<String>).add(item)
                                }
                            }
                        }
                    }
                    lists.add(cusRow)
                }
            }
        } catch (e: Exception) {
            e.printStackTrace()
            println("parseStringXml error: $e")
        }
        return lists
    }

    private fun getFolderName(file: File): String? {
        val path = file.absolutePath

        val paths =
            path.replace("\\", "/").split("/".toRegex()).dropLastWhile { it.isEmpty() }.toTypedArray()
        return if (paths.isNotEmpty()) {
            paths[paths.size - 1]
        } else null
    }

    private fun getFolderFileNameList(path: String): List<String> {
        val dir = File(path)
        val lists: MutableList<String> = ArrayList()
        if (dir.exists()) {
            val files = dir.listFiles()
            if (files != null) {
                val length = files.size
                for (i in 0..<length) {
                    val file = files[i]
                    val paths = file.absolutePath.toString()
                        .replace("\\", "/").split("/".toRegex()).dropLastWhile { it.isEmpty() }
                        .toTypedArray()
                    lists.add(paths[paths.size - 1])
                }
            }
        }
        return lists
    }

    private fun getIndexFromKey(key: String): Int {
        if (keyList.isNotEmpty()) {
            for (rowKey in keyList) {
                if (key == rowKey) {
                    return keyList.indexOf(key)
                }
            }
        }
        return -1
    }

    companion object {
        private const val ARRAY_TYPE_DECLARE = "type_array(do not modify)"
        private const val STRING_TYPE_DECLARE = "type_string(do not modify)"
        private const val STRING_NAME = "strings"
        private const val ARRAY_NAME = "arrays"
    }
}
