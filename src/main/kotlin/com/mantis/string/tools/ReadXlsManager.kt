package com.mantis.string.tools

import com.mantis.string.bean.XlsReadBean
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.*

class ReadXlsManager private constructor() : BaseSheet() {
    private var builder: XlsReadBean.Builder? = null
    private var isArrayFile = false
    private val map: MutableMap<Int, String?> = HashMap()
    private var stringKey: String? = null
    private var isLastItemString = false

    object Holder {
        var instance = ReadXlsManager()
    }

    fun readXls(builder2: XlsReadBean.Builder) {
        builder = builder2
        val file: File? = builder2.xlsFile?.let { File(builder2.rootPath, it) }
        if (file?.exists() == true) {
            try {
                val inputStream: InputStream = FileInputStream(file)
                val wb: Workbook = if (builder2.xlsFile?.toLowerCase()?.endsWith("xlsx") == true) {
                    XSSFWorkbook(inputStream)
                } else {
                    HSSFWorkbook(inputStream)
                }
                val sheetNum: Int = wb.numberOfSheets
                for (i in 0..<sheetNum) {
                    val sheet: Sheet = wb.getSheetAt(i) as Sheet
                    val firstRowIndex: Int = sheet.firstRowNum

                    MAX_ROW = getMaxRow(sheet) + 1
                    for (j in firstRowIndex..<MAX_ROW) {
                        val row: Row = sheet.getRow(j)
                        if (j == 0) {
                            readFirstRow(row)
                        } else {
                            readAllRows(row, j)
                        }
                    }
                }
            } catch (e: Exception) {
                // TODO: handle exception
                println("readXls, error: $e")
            }
        } else {
            println(("Cannot find " + builder2.rootPath) + File.separator + builder2.xlsFile)
        }
    }

    /**
     * get the max row of xls
     *
     * @param sheet
     * @return
     */
    private fun getMaxRow(sheet: Sheet): Int {
        val firstRow: Int = sheet.firstRowNum
        val lastRow: Int = sheet.lastRowNum
        var num = 0
        for (i in firstRow..<lastRow) {
            sheet.getRow(i)
            num++
        }
        return num
    }

    /**
     * get the first row
     * @param row
     */
    private fun readFirstRow(row: Row?) {
        if (row != null) {
            val firstCellIndex: Short = row.firstCellNum
            MAX_COLUMN = row.lastCellNum.toInt()

            try {
                for (j in firstCellIndex..<MAX_COLUMN) {
                    val cell: Cell = row.getCell(j)
                    var value: String
                    if (j == 0) {
                        value = cell.toString()
                        println("readFirstRow, value: $value")
                        if (value.contains("type_array")) {
                            isArrayFile = true
                        }
                    } else {
                        value = cell.toString()

                        val folderName = getFolderByLang(value)
                        println("language: $value , dir: $folderName")

                        val filepath =
                            createFolder(builder?.rootPath, builder?.fileFolderName + "/" + folderName)
                        map[j] = filepath

                        builder?.let { createFileAndData(map[j], isArrayFile, it) }
                    }

                }
            } catch (e: Exception) {
                //println("readFirstRow, exception: $e")
            }

        }
    }

    /**
     * read rows
     * @param row
     * @param rowIndex
     */
    private fun readAllRows(row: Row?, rowIndex: Int) {
        if (row != null) {
            val firstCellIndex: Short = row.firstCellNum
            val lastCellIndex: Short = row.lastCellNum
            for (cIndex in firstCellIndex..<lastCellIndex) {
                val cell: Cell = row.getCell(cIndex)
                if (isArrayFile) {
                    writeValueToArray(cIndex, cell, rowIndex)
                } else {
                    val path = map[cIndex]
                    writeValueToString(cIndex, cell, rowIndex, path)
                }
            }
        }
    }

    /**
     * write data to string
     */
    private fun writeValueToString(cIndex: Int, cell: Cell?, rowIndex: Int, dir: String?) {
        var value: String
        if (cell != null) {
            value = cell.toString()
            if (cIndex == 0) {
                stringKey = value
            } else {
                value = cell.toString().trim()
                value = value.replace("% ".toRegex(), "%")
                    .replace("1 ".toRegex(), "1").replace("2 ".toRegex(), "2")
                    .replace("3 ".toRegex(), "3").replace("4 ".toRegex(), "4")
                    .replace(" s".toRegex(), "s").replace(" d".toRegex(), "d")
                    .replace("％".toRegex(), "%").replace("\'", "\\'")

                val file = builder?.stringName?.let { File(dir, it) }
                if (file?.exists() == true) {
                    var fos: FileOutputStream? = null
                    try {
                        fos = FileOutputStream(file, true)
                        val builder3 = StringBuilder()
                        builder3.append("\t")
                            .append("<string name=\"").append(stringKey)
                            .append("\">").append(value).append("</string>")
                            .append("\r\n")
                        if (MAX_ROW - builder?.ignoreRow!! == rowIndex) {
                            builder3.append("</resources>\r\n")
                        }
                        fos.write(builder3.toString().toByteArray(charset("utf-8")))
                    } catch (e: Exception) {
                        e.printStackTrace()
                    } finally {
                        if (fos != null) {
                            try {
                                fos.close()
                            } catch (e: IOException) {
                                e.printStackTrace()
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * write data to array
     */
    private fun writeValueToArray(cIndex: Int, cell: Cell?, rowIndex: Int) {
        var value: String
        var sb: StringBuilder?
        if (cell != null) {
            value = cell.toString()
            value = value.replace(" ".toRegex(), "")

            if (cIndex == 0) {
                sb = StringBuilder()
                if (!isLastItemString) {
                    sb.append("\t")
                        .append("<string-array name =\"").append(value).append("\">")
                    writeDataToArrayFile(sb.toString(), null)
                } else {
                    sb = StringBuilder()
                    sb.append("\n")
                        .append("\t")
                        .append("</string-array>")
                    writeDataToArrayFile(sb.toString(), null)

                    sb = StringBuilder()
                    sb.append("\n\n")
                        .append("\t")
                        .append("<string-array name =\"").append(value).append("\">")
                    writeDataToArrayFile(sb.toString(), null)
                    isLastItemString = false
                }
            } else if (cIndex > 1) {
                sb = StringBuilder()
                sb.append("\n")
                    .append("\t\t")
                    .append("<item>")
                    .append(value)
                    .append("</item>")
                isLastItemString = true
                writeDataToArrayFile(sb.toString(), map[cIndex])
                if (rowIndex == MAX_ROW - builder?.ignoreRow!!
                    && cIndex == MAX_COLUMN - 1
                ) {
                    sb = StringBuilder()
                    sb.append("\n")
                        .append("\t")
                        .append("</string-array>\r\n")
                        .append("</resources>\r\n")
                    writeDataToArrayFile(sb.toString(), null)
                }
            }
        }
    }

    /**
     * write data to array
     * @param value
     * @param dir
     */
    private fun writeDataToArrayFile(value: String, dir: String?) {
        val files: MutableList<File> = ArrayList()
        if (dir != null) {
            val file = builder?.arrayName?.let { File(dir, it) }
            file?.let { files.add(it) }
        } else {
            for (i in 2..<MAX_COLUMN) {
                val path = map[i]
                val file = builder?.arrayName?.let { File(path, it) }
                file?.let { files.add(it) }
            }
        }
        for (file in files) {
            if (file.exists()) {
                var fos: FileOutputStream? = null
                try {
                    fos = FileOutputStream(file, true)
                    fos.write(value.toByteArray(charset("utf-8")))
                } catch (e: Exception) {
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
        }
    }

    companion object {
        private var MAX_COLUMN = 0
        private var MAX_ROW = 0
    }
}
