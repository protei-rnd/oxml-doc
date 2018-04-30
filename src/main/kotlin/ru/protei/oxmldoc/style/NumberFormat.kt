package ru.protei.oxmldoc.style

import java.util.stream.IntStream

class NumberFormat(val id: Int, var format: String){

    fun asNumberFormat(fractionDigits: Int): NumberFormat {
        return asNumberFormat(1, fractionDigits)
    }

    fun asNumberFormat(integerDigits: Int, fractionDigits: Int): NumberFormat {
        val sb = append(StringBuilder(), '0', integerDigits).append('.')
        format = append(sb, '0', fractionDigits).toString()
        return this
    }

    private fun append(sb: StringBuilder, char: Char, times: Int): StringBuilder {
        (0 .. times).forEach { sb.append(char) }
        return sb
    }
}