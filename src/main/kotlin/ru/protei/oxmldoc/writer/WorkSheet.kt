package ru.protei.oxmldoc.writer

import ru.protei.oxmldoc.common.ColumnRule
import ru.protei.oxmldoc.common.Row
import ru.protei.oxmldoc.common.RowsLimitRule.*
import java.io.IOException
import java.io.OutputStream
import java.io.UnsupportedEncodingException
import java.text.CharacterIterator
import java.text.SimpleDateFormat
import java.text.StringCharacterIterator
import java.util.*

class WorkSheet @Throws(IOException::class)
internal constructor(
        /**
         * @return the id
         */
        val id: Int,
        /**
         * @param name
         * the name to set
         */
        var name: String?,
        var out: OutputStream?, //  private File m_outFile;
        val writer: WorkBookWriter) {

    private var thisPageHeaderWritten = false

    private val colsRule = ArrayList<ColumnRule>()

    private var rowIndex: Long = 0

    /**
     * @return the partNumber
     */
    var partNumber = 1

    /**
     * @return the sourceName
     */
    /**
     * @param sourceName the sourceName to set
     */
    var sourceName: String? = null

    var columnRules: List<ColumnRule>
        get() = colsRule
        set(rules) {
            colsRule.clear()
            colsRule.addAll(rules)
        }

    init {
        val wsheetCnt = StringBuilder()
        wsheetCnt
                .append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
                .append(
                        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n")

        out!!.write(wsheetCnt.toString().toByteArray(charset("UTF-8")))
    }//    outFile = f;
    //new ZipOutputStream(new FileOutputStream (f));


    fun addColumnRule(r: ColumnRule) {
        colsRule.add(r)
    }

    @Throws(IOException::class)
    private fun writeHeaderIfDidnt() {
        if (!thisPageHeaderWritten) {
            val buf = StringBuilder()

            if (colsRule.size > 0) {

                buf.append("<cols>")

                for (r in colsRule) {
                    buf.append("<col min=\"").append(r.begin + 1).append("\"").append(
                            " max=\"").append(r.end + 1).append("\"")

                    if (r.width > 0)
                        buf.append(" width=\"").append(r.width).append("\"").append(
                                " customWidth=\"1\"")

                    if (r.style != null)
                        buf.append(" style=\"").append(r.style.id).append("\"")

                    buf.append("/>")
                }

                buf.append("</cols>")
            }

            buf.append("<sheetData>\n")
            out!!.write(buf.toString().toByteArray(charset("UTF-8")))
            thisPageHeaderWritten = true
        }
    }


    @Throws(IOException::class)
    fun appendRow(r: Row): WorkSheet {
        return if (rowIndex >= writer.getRowsLimit()) {
            when (writer.getRowsLimitRule()) {
                ERROR -> {
                    close()
                    throw IOException("Rows number per sheet limit exceeded")
                }

                SKIP -> this
            //        break;

                SPLIT -> {
                    val clone = writer.createSplittedClone(this)
                    close()
                    clone.appendRow(r)
                }

            // IGNORE
                IGNORE -> doAppendRow(r)
                else -> doAppendRow(r)
            }
        } else
            doAppendRow(r)
    }


    @Throws(IOException::class, UnsupportedEncodingException::class)
    protected fun doAppendRow(r: Row): WorkSheet {
        writeHeaderIfDidnt()

        rowIndex++
        val cnt = StringBuilder()

        cnt.append("<row r=\"").append(rowIndex).append("\"")

        when (r.rowStyle) {
            r.rowStyle.cellStyle != null ->
        }
        if (r.rowStyle != null) {
            if (r.rowStyle.cellStyle != null) {
                cnt.append(" s=\"").append(r.rowStyle.cellStyle.id).append("\"")
                        .append(" customFormat=\"1\"")
            }

            if (r.rowStyle.height > 0) {
                cnt.append(" ht=\"").append(r.rowStyle.height).append("\"").append(
                        " customHeight=\"1\"")
            }

            if (r.rowStyle.borderTop) {
                cnt.append(" thickTop=\"true\"")
            }

            if (r.rowStyle.borderBottom) {
                cnt.append(" thickBot=\"true\"")
            }
        }

        cnt.append(">")

        var cell = 0
        r.cells.forEach { c ->
            if (c.isEmpty) {
                cell++
                return@forEach
            }

            val cellStyle = c.style

            when (c.data) {
                is Number -> {
                    cnt.append("<c r=\"").append(columnNames[cell]).append(rowIndex).append("\"")
                    if (cellStyle != null)
                        cnt.append(" s=\"").append(cellStyle.id).append("\"")
                    cnt.append(">")

                    cnt.append("<v>").append(convertNumberToRawString(c.data)).append("</v>")
                    cnt.append("</c>")
                }
                is Date -> {
                    cnt.append("<c r=\"").append(columnNames[cell]).append(rowIndex)
                            .append("\"")
                    if (cellStyle != null)
                        cnt.append(" s=\"").append(cellStyle.id).append("\"")
                    cnt.append(" t=\"inlineStr\"")
                    cnt.append(">")

                    cnt.append("<is><t>").append(DATE_TIME_FMT.format(c.data))
                            .append("</t></is>")
                    cnt.append("</c>")
                }
                is Function<*> -> {
                    cnt.append("<c r=\"").append(columnNames[cell]).append(rowIndex)
                            .append("\"")
                    if (cellStyle != null)
                        cnt.append(" s=\"").append(cellStyle.id).append("\"")
                    cnt.append(" t=\"inlineStr\"")
                    cnt.append(">")

                    cnt.append("<f>").append(escapeXML((c.data).))
                            .append("</f>")
                    cnt.append("</c>")
                }
                else -> {
                    cnt.append("<c r=\"").append(columnNames[cell]).append(rowIndex)
                            .append("\" t=\"inlineStr\"")
                    if (cellStyle != null)
                        cnt.append(" s=\"").append(cellStyle.id).append("\"")
                    cnt.append(">")

                    cnt.append("<is><t>").append(escapeXML(c.data.toString())).append(
                            "</t></is>")
                    cnt.append("</c>")
                }
            }

            cell++
        }

        cnt.append("</row>\n")

        out!!.write(cnt.toString().toByteArray(charset("UTF-8")))

        return this
    }

    fun close() {
        if (out != null) {
            try {
                closeCurrentSheet()
            } catch (e: Throwable) {
                e.printStackTrace()
            } finally {
                try {
                    out!!.close()
                } catch (e: Throwable) {
                }

            }
        }

        out = null
    }

    @Throws(IOException::class)
    private fun closeCurrentSheet() {
        out!!.write("</sheetData>".toByteArray())
        out!!.write("</worksheet>".toByteArray())

        colsRule.clear()
    }

    companion object {
        private val alphabet = arrayOf(
                "A", "B", "C", "D", "E", "F",
                "G", "H", "I", "J", "K", "L",
                "M", "N", "O", "P", "Q", "R",
                "S", "T", "U", "V", "W", "X", "Y", "Z"
        )

        private val ASZ = alphabet.size

        private val columnNames = arrayOfNulls<String>(ASZ * (1 + ASZ * (1 + ASZ))) // x + x*x + x*x*x

        private val DATE_TIME_FMT = SimpleDateFormat("yyyy-MM-dd HH:mm:ss")

        init {
            var i = 0
            var j = ASZ
            var k = ASZ * (1 + ASZ)

            alphabet.forEach { str ->
                columnNames[i++] = str

                alphabet.forEach { secondStr ->
                    val tempStr = str + secondStr
                    columnNames[j++] = tempStr
                    alphabet.forEach { thirdStr -> columnNames[k++] = tempStr + thirdStr }
                }
            }
        }

        fun escapeXML(x: String): String {
            val result = StringBuilder()
            val iterator = StringCharacterIterator(x)
            var character = iterator.current()
            while (character != CharacterIterator.DONE) {
                when (character) {
                    '<' -> result.append("&lt;")
                    '>' -> result.append("&gt;")
                    '\"' -> result.append("&quot;")
                    '\'' -> result.append("&#039;")
                    '&' -> result.append("&amp;")
                    else -> //the char is not a special one
                        //add it to the result as is
                        result.append(character)
                }
                character = iterator.next()
            }
            return result.toString()
        }


        private fun convertNumberToRawString(number: Number?): String {
            val tx = number?.toString() ?: ""
            return tx.replace(',', '.')
        }
    }
}
