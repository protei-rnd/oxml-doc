package ru.protei.oxmldoc.writer;

import org.apache.log4j.Logger;
import ru.protei.oxmldoc.common.ColumnRule;
import ru.protei.oxmldoc.common.Row;
import ru.protei.oxmldoc.common.RowLimitRules;
import ru.protei.oxmldoc.style.CellStyle;
import ru.protei.oxmldoc.style.Fill;
import ru.protei.oxmldoc.style.Font;
import ru.protei.oxmldoc.style.NumberFormat;
import ru.protei.oxmldoc.system.SmartStreamConsumer;
import ru.protei.oxmldoc.system.SmartStreamFactory;
import ru.protei.oxmldoc.system.StreamConsumer;
import ru.protei.oxmldoc.system.StreamFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class WorkBookWriter implements AutoCloseable {
    /**
     * The maximum rows number per one sheet by default
     */
    private final static int ROWS_LIMIT_DEFAULT = 1000000;
    private static Logger log = Logger.getLogger(WorkBookWriter.class);
    private final String creator = "OpenXmlDoc-API";
    private final List<SheetDescriptor> sheets = new ArrayList<>();
    private final List<Font> fonts = new ArrayList<>();
    private final List<NumberFormat> numberFormats = new ArrayList<>();
    private final List<Fill> fills = new ArrayList<>();
    private final List<CellStyle> styles = new ArrayList<>();
    /**
     * The maximum rows number per one sheet for this instance
     */
    private int rowsLimit = ROWS_LIMIT_DEFAULT;
    private RowLimitRules rowsLimitRule = RowLimitRules.IGNORE;
    private String sheetNameTemplate = "%orig_%part";
    private StreamFactory streamFactory = SmartStreamFactory.DEFAULT;
    private ZipOutputStream zipOut;
    private WorkSheet currentSheet;
    private int numFormatStartId = 200;
    public WorkBookWriter() {
        this(ROWS_LIMIT_DEFAULT, RowLimitRules.IGNORE);
    }

    public WorkBookWriter(int rowsLimit, RowLimitRules rule) {
        this(rowsLimit, rule, SmartStreamFactory.DEFAULT);
    }

    public WorkBookWriter(int rowsLimit, RowLimitRules rule, SmartStreamFactory streamFactory) {
        this.streamFactory = streamFactory;
        this.rowsLimit = rowsLimit;
        this.rowsLimitRule = rule;
        final Font font = createFont();

        font.setFontSize("11");
        font.setColorTheme("1");
        font.setName("Calibri");
        font.setFamily(2);
        font.setCharset("1024");
        font.setScheme("minor");

        createCellStyle();
        createFill();
        createFill();
    }

    private static void pumpZipedFile2Stream(InputStream in, OutputStream out) {
        ZipInputStream zipIn = null;
        try {
            byte buf[] = new byte[4096];
            int read = 0;
            zipIn = new ZipInputStream(in);

            while (zipIn.getNextEntry() != null) {
                while ((read = zipIn.read(buf)) > 0) {
                    out.write(buf, 0, read);
                }
            }

            out.flush();
        } catch (Throwable ignored) {

        } finally {
            try {
                if (zipIn != null) {
                    zipIn.close();
                }
            } catch (Throwable ignored) {
            }
            ;
            try {
                in.close();
            } catch (Throwable ignored) {
            }
            ;
        }
    }

    private static ZipEntry putZipEntry(ZipOutputStream zipOut, String name,
                                        StringBuffer content) throws IOException {
        ZipEntry entry = new ZipEntry(name);
        zipOut.putNextEntry(entry);
        zipOut.write(content.toString().getBytes("UTF-8"));
        zipOut.closeEntry();
        return entry;
    }

    public void addColumnRule(ColumnRule rule) {
        currentSheet.addColumnRule(rule);
    }

    public void setColumnRules(List<ColumnRule> rules) {
        currentSheet.setColumnRules(rules);
    }

    public WorkBookWriter appendRow(Row row) throws IOException {
        currentSheet.appendRow(row);
        return this;
    }

    public WorkSheet getCurrentSheet() {
        return currentSheet;
    }

    public Fill createFill() {
        Fill f = new Fill(fills.size());
        fills.add(f);
        return f;
    }

    public Font createFont() {
        Font font = new Font(fonts.size());
        fonts.add(font);
        return font;
    }

    public CellStyle createCellStyle() {
        CellStyle st = new CellStyle(styles.size());
        styles.add(st);
        return st;
    }

    public NumberFormat createNumFormat() {
        NumberFormat fmt = new NumberFormat(numFormatStartId);
        numFormatStartId++;
        this.numberFormats.add(fmt);
        return fmt;
    }

    public WorkSheet getWorkSheet(String name) {
        WorkSheet sheet = null;

        // search last sheet with same name
        for (SheetDescriptor d : sheets)
            if (d.sheet.getSourceName() != null && d.sheet.getSourceName().equals(name) || d.sheet.getName().equals(name))
                sheet = d.sheet;

        return sheet;
    }

    private int findDescriptorIndex(WorkSheet sheet) {
        for (int id = 0; id < sheets.size(); id++) {
            if (sheets.get(id).sheet == sheet)
                return id;
        }

        return -1;
    }

    public WorkSheet createSplittedClone(WorkSheet sheet) throws IOException {

        final SheetDescriptor desc = new SheetDescriptor();
        desc.consumer = new SmartStreamConsumer();

        log.debug("create temp file for excell work sheet (clone): " + sheet.getName());

        ZipOutputStream zipOut = new ZipOutputStream(streamFactory.allocate(desc.consumer));

        zipOut.putNextEntry(new ZipEntry("excell_wsheet_tmp.xml"));

        int part = sheet.getPartNumber() + 1;
        String sourcename = sheet.getSourceName() == null ? sheet.getName() : sheet.getSourceName();

        String name = sheetNameTemplate.replace("%orig", sourcename);
        name = name.replace("%part", String.valueOf(part));

        desc.sheet = new WorkSheet("ws" + sheets.size() + 1 + "Id", name, zipOut, this);
        desc.sheet.setPartNumber(part);
        desc.sheet.setColumnRules(sheet.getColumnRules());
        desc.sheet.setSourceName(sourcename);

        currentSheet = desc.sheet;
        sheets.add(findDescriptorIndex(sheet) + 1, desc);
        return currentSheet;
    }

    public WorkSheet createNewSheet(String name) throws IOException {
        SheetDescriptor desc = new SheetDescriptor();
        desc.consumer = new SmartStreamConsumer();

        log.debug("create temp file for excell work sheet: " + name);

        ZipOutputStream zipOut = new ZipOutputStream(streamFactory.allocate(desc.consumer));
        zipOut.putNextEntry(new ZipEntry("excell_wsheet_tmp.xml"));

        desc.sheet = new WorkSheet("ws" + sheets.size() + 1 + "Id", name, zipOut, this);

        currentSheet = desc.sheet;
        sheets.add(desc);
        return currentSheet;
    }

    private void doWrite(ZipOutputStream zipOut) throws IOException {
        int c = 1;
        for (SheetDescriptor d : sheets) {
            d.sheet.close();
            zipOut.putNextEntry(new ZipEntry("xl/worksheets/sheet" + c + ".xml"));
            InputStream rawIn = d.consumer.createInput();
            if (rawIn != null)
                pumpZipedFile2Stream(rawIn, zipOut);
            c++;
        }


        // build doc-tables
        StringBuffer contentDescriptor = new StringBuffer();

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
                .append(
                        "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>\n")
                .append(
                        "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\n");

        for (int i = 0; i < sheets.size(); i++) {
            contentDescriptor
                    .append("<Override PartName=\"/xl/worksheets/sheet")
                    .append(i + 1)
                    .append(
                            ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />\n");
        }

        contentDescriptor.append("</Types>");

        StringBuffer rootRels = new StringBuffer();
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
                .append("</Relationships>");

        StringBuffer appPropCnt = new StringBuffer();
        appPropCnt
                .append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
                .append(
                        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">\n")
                .append("<Application>Open XML Document API</Application>\n")
                .append("</Properties>");

        StringBuffer corePropCnt = new StringBuffer();
        corePropCnt
                .append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
                .append(
                        "<coreProperties xmlns=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
                                + "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" "
                                + "xmlns:dcterms=\"http://purl.org/dc/terms/\" "
                                + "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" "
                                + "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n"
                                + "<dc:creator>Open XML Document API</dc:creator>\n"
                                + "</coreProperties>");

        StringBuffer wbookContent = new StringBuffer();
        wbookContent
                .append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
                .append(
                        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n")
                .append(
                        "<fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/>\n")
                .append("<sheets>\n");

        for (int i = 0; i < sheets.size(); i++) {
            SheetDescriptor desc = sheets.get(i);
            wbookContent.append("<sheet name=\"").append(WorkSheet.escapeXML(desc.sheet.getName())).append(
                    "\" sheetId=\"").append(i + 1).append("\" r:id=\"")
                    .append(desc.sheet.getId()).append("\"/>\n");
        }

        wbookContent.append("</sheets>\n").append("</workbook>");

        StringBuffer wbookRels = new StringBuffer();

        wbookRels
                .append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
                .append(
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");

        for (int i = 0; i < sheets.size(); i++) {
            SheetDescriptor desc = sheets.get(i);
            wbookRels
                    .append("<Relationship Id=\"")
                    .append(desc.sheet.getId())
                    .append(
                            "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\"")
                    .append(" Target=\"worksheets/sheet").append(i + 1).append(
                    ".xml\" />\n");
        }

        wbookRels
                .append(
                        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n")
                .append("</Relationships>");

        StringBuffer stylesCnt = new StringBuffer();
        stylesCnt
                .append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
                .append(
                        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

        stylesCnt.append("<numFmts count=\"").append(numberFormats.size()).append("\">");
        for (NumberFormat fmt : numberFormats) {
            stylesCnt.append("<numFmt numFmtId=\"").append(fmt.getId()).append("\"");
            stylesCnt.append(" formatCode=\"").append(WorkSheet.escapeXML(fmt.getFormat())).append("\"");
            stylesCnt.append("/>");
        }
        stylesCnt.append("</numFmts>");

        stylesCnt.append("<fonts count=\"").append(fonts.size()).append("\">");
        for (Font f : fonts) {
            stylesCnt.append("<font>");

            if (f.getFontSize() != null)
                stylesCnt.append("<sz val=\"").append(f.getFontSize()).append("\"/>");

            if (f.getColor() != null)
                stylesCnt.append("<color rgb=\"").append(WorkSheet.escapeXML(f.getColor())).append("\"/>");
            else if (f.getColorTheme() != null)
                stylesCnt.append("<color theme=\"").append(WorkSheet.escapeXML(f.getColorTheme())).append("\"/>");

            stylesCnt.append("<name val=\"").append(WorkSheet.escapeXML((f.getName() != null ? f.getName() : "Arial"))).append("\"/>");

            if (f.getFamily() != 0)
                stylesCnt.append("<family val=\"").append(f.getFamily()).append("\"/>");

            if (f.getCharset() != null)
                stylesCnt.append("<charset val=\"").append(WorkSheet.escapeXML(f.getCharset())).append("\"/>");

            if (f.getScheme() != null)
                stylesCnt.append("<scheme val=\"").append(WorkSheet.escapeXML(f.getScheme())).append("\"/>");

            if (f.isBold())
                stylesCnt.append("<b/>");

            if (f.isItalic())
                stylesCnt.append("<i/>");

            if (f.isUnderlined())
                stylesCnt.append("<u/>");

            if (f.isStriked())
                stylesCnt.append("<strike/>");

            stylesCnt.append("</font>");
        }
        stylesCnt.append("</fonts>");

        stylesCnt.append("<fills count=\"").append(fills.size()).append("\">");
        for (Fill f : fills) {
            stylesCnt.append("<fill><patternFill patternType=\"").append(f.getPattern())
                    .append("\"");

            if (f.getBackgroundColor() == null && f.getFrontColor() == null) {
                stylesCnt.append("/>");
            } else {
                stylesCnt.append(">");

                if (f.getFrontColor() != null) {
                    stylesCnt.append("<fgColor rgb=\"").append(
                            Integer.toHexString(f.getFrontColor().getRgb()).toUpperCase())
                            .append("\"/>");
                }

                if (f.getBackgroundColor() != null) {
                    stylesCnt.append("<bgColor rgb=\"").append(
                            Integer.toHexString(f.getBackgroundColor().getRgb()).toUpperCase())
                            .append("\"/>");
                }

                stylesCnt.append("</patternFill>");
            }
            stylesCnt.append("</fill>");
        }

        stylesCnt.append("</fills>");

        stylesCnt
                .append("<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>");

        stylesCnt.append("<cellStyleXfs count=\"1\">");
        stylesCnt
                .append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>");
        stylesCnt.append("</cellStyleXfs>");

        stylesCnt.append("<cellXfs count=\"").append(styles.size())
                .append("\">");
        for (CellStyle s : styles) {
            if (s.getId() == 0) {
                stylesCnt
                        .append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
                continue;
            }

            stylesCnt.append("<xf");

            stylesCnt.append(" numFmtId=\"").append(s.getNumberFormat() != null ? s.getNumberFormat().getId() : 0).append("\"");
            stylesCnt.append(" borderId=\"").append(0).append("\"");

            if (s.getAligment() != null) {
                stylesCnt.append(" applyAlignment=\"true\"");
            }

            stylesCnt.append(" fillId=\"").append(
                    (s.getFill() != null ? s.getFill().getId() : 0)).append("\"");
            if (s.getFill() != null)
                stylesCnt.append(" applyFill=\"true\"");

            if (s.getFont() != null) {
                stylesCnt.append(" applyFont=\"true\"");
                stylesCnt.append(" fontId=\"").append(s.getFont().getFontId()).append(
                        "\"");
            } else {
                stylesCnt.append(" fontId=\"").append(0).append("\"");
            }

            stylesCnt.append(" xfId=\"").append(0).append("\"");

            stylesCnt.append(">");

            if (s.getAligment() != null) {
                stylesCnt.append("<alignment");
                if (s.getAligment().getHorizAlign() != null)
                    stylesCnt.append(" horizontal=\"").append(
                            s.getAligment().getHorizAlign().name()).append("\"");

                if (s.getAligment().getVertAlign() != null)
                    stylesCnt.append(" vertical=\"")
                            .append(s.getAligment().getVertAlign().name()).append("\"");

                if (s.getAligment().getWrapText() != null)
                    stylesCnt.append(" wrapText=\"").append(s.getAligment().getWrapText())
                            .append("\"");

                stylesCnt.append("/>");
            }

            stylesCnt.append("</xf>");
        }
        stylesCnt.append("</cellXfs>");

        stylesCnt.append("</styleSheet>");

        putZipEntry(zipOut, "[Content_Types].xml", contentDescriptor);
        putZipEntry(zipOut, "_rels/.rels", rootRels);

        putZipEntry(zipOut, "docProps/app.xml", appPropCnt);
        putZipEntry(zipOut, "docProps/core.xml", corePropCnt);

        putZipEntry(zipOut, "xl/_rels/workbook.xml.rels", wbookRels);
        putZipEntry(zipOut, "xl/workbook.xml", wbookContent);
        putZipEntry(zipOut, "xl/styles.xml", stylesCnt);

        zipOut.flush();
        zipOut.close();
    }


    public void writeTo(OutputStream out) throws IOException {
        try (ZipOutputStream zout = new ZipOutputStream(out)) {
            doWrite(zout);
        }
    }

    public void writeTo(File f) throws IOException {
        try (OutputStream out = new FileOutputStream(f, false)) {
            writeTo(out);
        }
    }

    @Override
    public void close() throws IOException {
        if (zipOut != null) {
            doWrite(zipOut);
            zipOut = null;
        }

        for (SheetDescriptor d : sheets) {
            d.sheet.close();
            d.consumer.release();
        }

        this.sheets.clear();
        this.currentSheet = null;
        this.styles.clear();
        this.fills.clear();
        this.fonts.clear();
    }

    /**
     * @return the creator
     */
    public String getCreator() {
        return creator;
    }

    /**
     * @return the rowsLimit
     */
    public int getRowsLimit() {
        return rowsLimit;
    }

    /**
     * @return the rowsLimitRule
     */
    public RowLimitRules getRowLimitRules() {
        return rowsLimitRule;
    }

    private class SheetDescriptor {
        WorkSheet sheet;
        StreamConsumer consumer;
    }
}
