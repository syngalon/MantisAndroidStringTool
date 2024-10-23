package com.mantis.string.bean

class StringRow {
    var key: String? = null
    var value: String? = null
    private var isArray = false
    var items: List<String>? = null
    override fun toString(): String {
        return ("CusRow [key=" + key + ", value=" + value + ", isArray="
                + isArray + ", items=" + items + "]")
    }
}