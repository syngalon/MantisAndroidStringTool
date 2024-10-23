package com.mantis.string.bean

class XlsWriteBean(val builder: Builder) {

    class Builder {
        var rootPath: String? = null
        var xlsName = "workbook.xls"
        var fileFolderName: String? = null
        fun setRootPath(rootPath: String?): Builder {
            this.rootPath = rootPath
            return this
        }

        fun setXlsName(xlsName: String): Builder {
            this.xlsName = xlsName
            return this
        }

        fun setFileFolderName(fileFolderName: String?): Builder {
            this.fileFolderName = fileFolderName
            return this
        }

        fun builder(): XlsWriteBean {
            return XlsWriteBean(this)
        }
    }

    companion object {
        private fun checkNull(builder: Builder) {
            if (builder.rootPath == null) {
                throw NullPointerException("you need to set root path!")
            }
            if (builder.fileFolderName == null) {
                throw NullPointerException("you need to set file name!")
            }
        }
    }
}
