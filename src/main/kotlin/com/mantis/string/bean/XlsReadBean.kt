package com.mantis.string.bean

class XlsReadBean(val builder: Builder) {

    class Builder {
        var rootPath: String? = null
        var xlsFile: String? = null
        var fileFolderName: String? = null
        var ignoreRow = 0
        var stringName = "strings.xml"
        var arrayName = "array.xml"
        fun setRootPath(rootPath: String?): Builder {
            this.rootPath = rootPath
            return this
        }

        fun setFileFolderName(fileFolderName: String?): Builder {
            this.fileFolderName = fileFolderName
            return this
        }

        fun setXlsFile(xlsFile: String?): Builder {
            this.xlsFile = xlsFile
            return this
        }

        fun setIgnoreRow(ignoreRow: Int): Builder {
            this.ignoreRow = ignoreRow
            return this
        }

        fun setStringName(stringName: String): Builder {
            this.stringName = stringName
            return this
        }

        fun setArrayName(arrayName: String): Builder {
            this.arrayName = arrayName
            return this
        }

        fun builder(): XlsReadBean {
            checkNull(this)
            return XlsReadBean(this)
        }
    }

    companion object {
        private fun checkNull(builder: Builder) {
            if (builder.rootPath == null) {
                throw NullPointerException("please set the root path!")
            }
            if (builder.xlsFile == null) {
                throw NullPointerException("please set xls file file!")
            }
            if (builder.fileFolderName == null) {
                throw NullPointerException("please set file name!")
            }
        }
    }
}
