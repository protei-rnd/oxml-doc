package ru.protei.oxmldoc.writer

import ru.protei.oxmldoc.common.Row
import ru.protei.oxmldoc.common.RowsLimitRule
import ru.protei.oxmldoc.common.RowsLimitRule.*
import ru.protei.oxmldoc.style.CellStyle
import ru.protei.oxmldoc.style.Fill
import ru.protei.oxmldoc.style.Font
import ru.protei.oxmldoc.style.NumberFormat
import ru.protei.oxmldoc.system.SmartStreamConsumer
import ru.protei.oxmldoc.system.SmartStreamFactory
import ru.protei.oxmldoc.system.StreamConsumer
import ru.protei.oxmldoc.system.StreamFactory
import java.io.*
import java.util.ArrayList
import java.util.zip.ZipEntry
import java.util.zip.ZipInputStream
import java.util.zip.ZipOutputStream

class WorkBookWriter constructor(val rowsLimit: Int, val rowsLimitRule: RowsLimitRule, val streamFactory: SmartStreamFactory,
                                 val font: Font, val cellStyle: CellStyle, val fill: Fill) {

    /**
     * The maximum rows number per one sheet for this instance
     */

    /**
     *
     */

    private val sheetNameTemplate = "%orig_%part"

    private var zipOut: ZipOutputStream? = null

    private var sheets: MutableList<SheetDesc> = ArrayList()

    private var currentSheet: WorkSheet? = null

    /**
     * @return the creator
     */
    /**
     * @param creator the creator to set
     */
    private var creator = "OpenXmlDoc-API"

    private var fontList = ArrayList<Font>()

    private val numberFormats = ArrayList<NumberFormat>()

    private var fillList = ArrayList<Fill>()

    private var cellStylesList = ArrayList<CellStyle>()


    private var numFormatStartId = 200

    internal inner class SheetDesc {
        var sheet: WorkSheet? = null
        var consumer: StreamConsumer = SmartStreamConsumer()
    }

    constructor(rowsLimit: Int = ROWS_LIMIT_DEFAULT, rule: RowsLimitRule = RowsLimitRule.IGNORE)
            : this(rowsLimit, rule, SmartStreamFactory.DEFAULT) {}

    constructor(rowsLimit: Int, rowsLimitRule: RowsLimitRule, streamFactory: SmartStreamFactory)
            : this(rowsLimit, ){
        val f = createFont()

        f.fontSize = "11"
        f.colorTheme = "1"
        f.name = "Calibri"
        f.family = 2
        f.charset = "204"
        f.scheme = "minor"

        createCellStyle()
        createFill()
        createFill()
    }

    fun addColumnRule(r: Tm_ColumnRule) {
        currentSheet!!.addColumnRule(r)
    }

    fun setColumnRules(rules: List<Tm_ColumnRule>) {
        currentSheet!!.setColumnRules(rules)
    }


    @Throws(IOException::class)
    fun appendRow(r: Row): Tm_WorkBookWriter {
        /**
         * in case of auto-split reference would be replaced
         */
        currentSheet!!.appendRow(r)
        return this
    }

    fun createFill(): Fill {
        val f = Fill(m_FillList!!.size)
        m_FillList!!.add(f)
        return f
    }

    fun createFont(): Font {
        val f = Font(m_FontList!!.size)
        m_FontList!!.add(f)
        return f
    }

    fun createCellStyle(): CellStyle {
        val st = CellStyle(m_CellStylesList!!.size)
        m_CellStylesList!!.add(st)
        return st
    }

    fun createNumFormat(): NumFormat {
        val fmt = NumFormat(numFormatStartId)
        numFormatStartId++
        this.m_NumFormats.add(fmt)
        return fmt
    }


    fun getWorkSheet(name: String): Tm_WorkSheet? {
        var sheet: Tm_WorkSheet? = null

        // search last sheet with same name
        for (d in m_Sheets!!)
            if (d.sheet!!.getSourceName() != null && d.sheet!!.getSourceName().equals(name) || d.sheet!!.getName().equals(name))
                sheet = d.sheet

        return sheet
    }

    private fun findDescriptorIndex(sheet: Tm_WorkSheet): Int {
        for (idx in m_Sheets!!.indices) {
            if (m_Sheets!![idx].sheet === sheet)
                return idx
        }

        return -1
    }

    @Throws(IOException::class)
    internal fun createSplittedClone(sheet: Tm_WorkSheet): Tm_WorkSheet {

        val desc = Tm_SheetDesc()
        desc.consumer = StreamConsumer()//tempFile = File.createTempFile("excell", "wb.cnt.zip");

        log.debug("create temp file for excell work sheet (clone): " + sheet.getName())

        val zipOut = ZipOutputStream(streamFactory.allocate(desc.consumer))

        zipOut.putNextEntry(ZipEntry("excell_wsheet_tmp.xml"))

        val part = sheet.getPartNumber() + 1
        val sourcename = if (sheet.getSourceName() == null) sheet.getName() else sheet.getSourceName()

        var name = sheetNameTemplate.replace("%orig", sourcename)
        name = name.replace("%part", part.toString())

        desc.sheet = Tm_WorkSheet("ws" + m_Sheets!!.size + 1 + "Id", name, zipOut, this)
        desc.sheet!!.setPartNumber(part)
        desc.sheet!!.setColumnRules(sheet.getColumnRules())
        desc.sheet!!.setSourceName(sourcename)

        currentSheet = desc.sheet
        m_Sheets!!.add(findDescriptorIndex(sheet) + 1, desc)

        //    m_ZipOut.putNextEntry(new ZipEntry("xl/worksheets/sheet" + m_Sheets.size()
        //        + ".xml"));
        return currentSheet
    }

    @Throws(IOException::class)
    fun createNewSheet(name: String): WorkSheet {
        val desc = SheetDesc()
        desc.consumer = SmartStreamConsumer()

        log.debug("create temp file for excell work sheet: $name")

        val zipOut = ZipOutputStream(streamFactory.allocate(desc.consumer))
        zipOut.putNextEntry(ZipEntry("excell_wsheet_tmp.xml"))

        desc.sheet = Tm_WorkSheet("ws" + m_Sheets!!.size + 1 + "Id", name, zipOut, this)

        currentSheet = desc.sheet
        m_Sheets!!.add(desc)

        //    m_ZipOut.putNextEntry(new ZipEntry("xl/worksheets/sheet" + m_Sheets.size()
        //        + ".xml"));
        return currentSheet
    }

    @Throws(IOException::class)
    protected fun doWrite(zipOut: ZipOutputStream) {
        var c = 1
        for (d in m_Sheets!!) {
            d.sheet!!.close()
            zipOut.putNextEntry(ZipEntry("xl/worksheets/sheet$c.xml"))
            val rawIn = d.consumer!!.createInput()
            if (rawIn != null)
                pumpZipedFile2Stream(rawIn, zipOut)
            //	      d.tempFile.delete();
            c++
        }


        // build doc-tables
        val contentDescriptor = StringBuffer()

        contentDescriptor
                .append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
                .append(
                        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n")
                .append(
                        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>\n")
                .append(
                        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n")
                .append(
                        "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n")
                .append(
                        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" />\n")
                // .append("<Override PartName=\"/xl/sharedStrings.xml\"
                // ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>\n")
                .append(
                        "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>\n")
                .append(
                        "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\n")

        for (i in m_Sheets!!.indices) {
            contentDescriptor
                    .append("<Override PartName=\"/xl/worksheets/sheet")
                    .append(i + 1)
                    .append(
                            ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />\n")
        }

        contentDescriptor.append("</Types>")

        val rootRels = StringBuffer()
        rootRels
                .append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
                .append(
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n")
                .append(
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n")
                .append(
                        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>\n")
                .append(
                        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>\n")
                .append("</Relationships>")

        val appPropCnt = StringBuffer()
        appPropCnt
                .append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
                .append(
                        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">\n")
                .append("<Application>Open XML Document API</Application>\n")
                .append("</Properties>")

        val corePropCnt = StringBuffer()
        corePropCnt
                .append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
                .append(
                        "<coreProperties xmlns=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
                                + "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" "
                                + "xmlns:dcterms=\"http://purl.org/dc/terms/\" "
                                + "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" "
                                + "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n"
                                + "<dc:creator>Open XML Document API</dc:creator>\n" +
                                // "<dcterms:created
                                // xsi:type=\"dcterms:W3CDTF\">2010-08-15T06:50:14Z</dcterms:created>\n"
                                // +
                                "</coreProperties>")

        val wbookContent = StringBuffer()
        wbookContent
                .append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
                .append(
                        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n")
                .append(
                        "<fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/>\n")
                // .append("<workbookPr dateCompatibility=\"0\" />\n")
                .append("<sheets>\n")

        for (i in m_Sheets!!.indices) {
            val desc = m_Sheets!![i]
            wbookContent.append("<sheet name=\"").append(Tm_WorkSheet.escapeXML(desc.sheet!!.getName())).append(
                    "\" sheetId=\"").append(i + 1).append("\" r:id=\"")
                    .append(desc.sheet!!.getId()).append("\"/>\n")
        }

        wbookContent.append("</sheets>\n").append("</workbook>")

        val wbookRels = StringBuffer()

        wbookRels
                .append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
                .append(
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n")

        for (i in m_Sheets!!.indices) {
            val desc = m_Sheets!![i]
            wbookRels
                    .append("<Relationship Id=\"")
                    .append(desc.sheet!!.getId())
                    .append(
                            "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\"")
                    .append(" Target=\"worksheets/sheet").append(i + 1).append(
                            ".xml\" />\n")
        }
        // .append("<Relationship Id=\"rId2\"
        // Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\"
        // Target=\"sharedStrings.xml\"/>\n")
        wbookRels
                .append(
                        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n")
                .append("</Relationships>")

        val stylesCnt = StringBuffer()
        stylesCnt
                .append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
                .append(
                        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">")

        stylesCnt.append("<numFmts count=\"").append(m_NumFormats.size).append("\">")
        for (fmt in m_NumFormats) {
            stylesCnt.append("<numFmt numFmtId=\"").append(fmt.getId()).append("\"")
            stylesCnt.append(" formatCode=\"").append(Tm_WorkSheet.escapeXML(fmt.getFormat())).append("\"")
            stylesCnt.append("/>")
        }
        stylesCnt.append("</numFmts>")

        stylesCnt.append("<fonts count=\"").append(m_FontList!!.size).append("\">")
        for (f in m_FontList!!) {
            stylesCnt.append("<font>")

            if (f.fontSize != null)
                stylesCnt.append("<sz val=\"").append(f.fontSize).append("\"/>")

            if (f.color != null)
                stylesCnt.append("<color rgb=\"" + Tm_WorkSheet.escapeXML(f.color) + "\"/>")
            else if (f.colorTheme != null)
                stylesCnt.append("<color theme=\"").append(Tm_WorkSheet.escapeXML(f.colorTheme)).append("\"/>")

            stylesCnt.append("<name val=\"").append(Tm_WorkSheet.escapeXML((if (f.name != null) f.name else "Arial"))).append("\"/>")

            if (f.family != null)
                stylesCnt.append("<family val=\"" + f.family + "\"/>")

            if (f.charset != null)
                stylesCnt.append("<charset val=\"" + Tm_WorkSheet.escapeXML(f.charset) + "\"/>")

            if (f.scheme != null)
                stylesCnt.append("<scheme val=\"" + Tm_WorkSheet.escapeXML(f.scheme) + "\"/>")

            if (f.bold)
                stylesCnt.append("<b/>")

            if (f.italic)
                stylesCnt.append("<i/>")

            if (f.underline)
                stylesCnt.append("<u/>")

            if (f.strike)
                stylesCnt.append("<strike/>")

            stylesCnt.append("</font>")
        }
        stylesCnt.append("</fonts>")

        stylesCnt.append("<fills count=\"").append(m_FillList!!.size).append("\">")
        for (f in m_FillList!!) {
            stylesCnt.append("<fill><patternFill patternType=\"").append(f.pattern)
                    .append("\"")

            if (f.m_BgColor == null && f.m_FgColor == null) {
                stylesCnt.append("/>")
            } else {
                stylesCnt.append(">")

                if (f.m_FgColor != null) {
                    stylesCnt.append("<fgColor rgb=\"").append(
                            Integer.toHexString(f.m_FgColor.rgb).toUpperCase())
                            .append("\"/>")
                }

                if (f.m_BgColor != null) {
                    stylesCnt.append("<bgColor rgb=\"").append(
                            Integer.toHexString(f.m_BgColor.rgb).toUpperCase())
                            .append("\"/>")
                }

                stylesCnt.append("</patternFill>")
            }
            stylesCnt.append("</fill>")
        }

        stylesCnt.append("</fills>")

        stylesCnt
                .append("<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>")

        stylesCnt.append("<cellStyleXfs count=\"1\">")
        stylesCnt
                .append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>")
        stylesCnt.append("</cellStyleXfs>")

        stylesCnt.append("<cellXfs count=\"").append(m_CellStylesList!!.size)
                .append("\">")
        for (s in m_CellStylesList!!) {
            if (s.getIdx() === 0) {
                stylesCnt
                        .append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>")
                continue
            }

            stylesCnt.append("<xf")

            stylesCnt.append(" numFmtId=\"").append(if (s.numFormat != null) s.numFormat.getId() else 0).append("\"")
            stylesCnt.append(" borderId=\"").append(0).append("\"")

            if (s.m_Aligment != null) {
                stylesCnt.append(" applyAlignment=\"true\"")
            }

            stylesCnt.append(" fillId=\"").append(
                    (if (s.m_Fill != null) s.m_Fill.getIdx() else 0)).append("\"")
            if (s.m_Fill != null)
                stylesCnt.append(" applyFill=\"true\"")

            if (s.m_Font != null) {
                stylesCnt.append(" applyFont=\"true\"")
                stylesCnt.append(" fontId=\"").append(s.m_Font.getFontIdx()).append(
                        "\"")
            } else {
                stylesCnt.append(" fontId=\"").append(0).append("\"")
            }

            stylesCnt.append(" xfId=\"").append(0).append("\"")

            stylesCnt.append(">")

            if (s.m_Aligment != null) {
                stylesCnt.append("<alignment")
                if (s.m_Aligment.horizAlign != null)
                    stylesCnt.append(" horizontal=\"").append(
                            s.m_Aligment.horizAlign.name()).append("\"")

                if (s.m_Aligment.vertAlign != null)
                    stylesCnt.append(" vertical=\"")
                            .append(s.m_Aligment.vertAlign.name()).append("\"")

                if (s.m_Aligment.wrapText != null)
                    stylesCnt.append(" wrapText=\"").append(s.m_Aligment.wrapText)
                            .append("\"")

                stylesCnt.append("/>")
            }

            stylesCnt.append("</xf>")
        }
        stylesCnt.append("</cellXfs>")

        stylesCnt.append("</styleSheet>")

        putZipEntry(zipOut, "[Content_Types].xml", contentDescriptor)
        putZipEntry(zipOut, "_rels/.rels", rootRels)

        putZipEntry(zipOut, "docProps/app.xml", appPropCnt)
        putZipEntry(zipOut, "docProps/core.xml", corePropCnt)

        putZipEntry(zipOut, "xl/_rels/workbook.xml.rels", wbookRels)
        putZipEntry(zipOut, "xl/workbook.xml", wbookContent)
        // putZipEntry(m_ZipOut, "xl/sharedStrings.xml", sharedStrCnt);
        putZipEntry(zipOut, "xl/styles.xml", stylesCnt)

        zipOut.flush()
        zipOut.close()
        //	    m_ZipOut.close();
    }


    @Throws(IOException::class)
    fun writeTo(out: OutputStream) {

        val zout = ZipOutputStream(out)

        try {
            doWrite(zout)
        } finally {
            try {
                zout.close()
            } catch (e: Throwable) {
            }

        }
    }

    @Throws(IOException::class)
    fun writeTo(f: File) {
        val out = FileOutputStream(f, false)
        try {
            writeTo(out)
        } finally {
            if (out != null)
                try {
                    out.close()
                } catch (e: Throwable) {
                }

        }
    }

    @Throws(IOException::class)
    fun close() {
        if (m_ZipOut != null) {
            doWrite(m_ZipOut)
            m_ZipOut = null
        }

        for (d in m_Sheets!!) {
            d.sheet!!.close()
            d.consumer!!.release()
            //if (d.tempFile.exists())
            //	  d.tempFile.delete();
        }

        this.m_Sheets!!.clear()
        this.currentSheet = null
        this.m_Sheets = null
        this.m_CellStylesList!!.clear()
        this.m_FillList!!.clear()
        this.m_FontList!!.clear()
        this.m_CellStylesList = null
        this.m_FillList = null
        this.m_FontList = null
    }

    companion object {
        // private WeakReference<OutputStream> m_Out;
        private val log = Logger.getLogger(Tm_WorkBookWriter::class.java)

        /**
         * The maximum rows number per one sheet by default
         */
        var ROWS_LIMIT_DEFAULT = 1000000


        private fun pumpZipedFile2Stream(`in`: InputStream, out: OutputStream) {
            //    FileInputStream in = null;
            var zipIn: ZipInputStream? = null
            try {
                val buf = ByteArray(4096)
                var read = 0
                zipIn = ZipInputStream(`in`)

                while (zipIn!!.nextEntry != null) {
                    while ((read = zipIn!!.read(buf)) > 0) {
                        out.write(buf, 0, read)
                    }
                }

                out.flush()
            } catch (ignored: Throwable) {

            } finally {
                try {
                    if (zipIn != null) {
                        zipIn!!.close()
                    }
                } catch (ignored: Throwable) {
                }

                try {
                    `in`.close()
                } catch (ignored: Throwable) {
                }

            }
        }

        @Throws(IOException::class)
        private fun putZipEntry(zipOut: ZipOutputStream, name: String,
                                content: StringBuffer): ZipEntry {
            val entry = ZipEntry(name)
            zipOut.putNextEntry(entry)
            zipOut.write(content.toString().toByteArray(charset("UTF-8")))
            zipOut.closeEntry()
            return entry
        }
    }
}
