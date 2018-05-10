package ru.protei.oxmldoc.writer;

import ru.protei.oxmldoc.common.Cell;
import ru.protei.oxmldoc.common.ColumnRule;
import ru.protei.oxmldoc.common.Row;
import ru.protei.oxmldoc.style.CellStyle;

import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.text.CharacterIterator;
import java.text.SimpleDateFormat;
import java.text.StringCharacterIterator;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import java.util.function.Supplier;

public class WorkSheet implements AutoCloseable {
    private static final String[] alphabet =
            {
                    "A", "B", "C", "D", "E",
                    "F", "G", "H", "I", "J",
                    "K", "L", "M", "N", "O",
                    "P", "Q", "R", "S", "T",
                    "U", "V", "W", "X", "Y",
                    "Z"
            };

    private static final int ASZ = alphabet.length;

    private static String[] columnNames = new String[ASZ * (1 + ASZ * (1 + ASZ))]; // x + x*x + x*x*x

    private static SimpleDateFormat DATE_TIME_FMT = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    static {
        int i = 0, j = ASZ, k = ASZ * (1 + ASZ);

        for (String str_1 : alphabet) {
            columnNames[i++] = str_1;

            for (String str_2 : alphabet) {
                String ts_1 = str_1 + str_2;
                columnNames[j++] = ts_1;

                for (String str_3 : alphabet)
                    columnNames[k++] = ts_1 + str_3;
            }
        }
    }

    private String sheetName;

    private String id;

    private boolean isHeaderWritten = false;

    private List<ColumnRule> columnRules = new ArrayList<>();

    private OutputStream outputStream;

    private long rowIndex = 0;

    private WorkBookWriter writer;

    private int partNumber = 1;

    private String sourceName;

    public WorkSheet(String id, String name, OutputStream out, WorkBookWriter writer) throws IOException {
        this.writer = writer;

        sheetName = name;
        this.id = id;
        outputStream = out;

        StringBuilder wsheetCnt = new StringBuilder();
        wsheetCnt
                .append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
                .append(
                        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n");

        outputStream.write(wsheetCnt.toString().getBytes("UTF-8"));
    }


    public void addColumnRule(ColumnRule r) {
        columnRules.add(r);
    }

    public void setColumnRules(List<ColumnRule> rules) {
        columnRules.clear();
        columnRules.addAll(rules);
    }

    public List<ColumnRule> getColumnRules() {
        return Collections.unmodifiableList(columnRules);
    }

    /**
     * @return the name
     */
    public String getName() {
        return sheetName;
    }

    /**
     * @param name the name to set
     */
    public void setName(String name) {
        sheetName = name;
    }

    /**
     * @return the id
     */
    public String getId() {
        return id;
    }

    private void writeHeaderIfDidnt() throws IOException {
        if (!isHeaderWritten) {
            StringBuilder buf = new StringBuilder();

            if (columnRules.size() > 0) {

                buf.append("<cols>");

                for (ColumnRule r : columnRules) {
                    buf.append("<col min=\"").append(r.getBegin() + 1).append("\"").append(
                            " max=\"").append(r.getEnd() + 1).append("\"");

                    if (r.getWidth() > 0)
                        buf.append(" width=\"").append(r.getWidth()).append("\"").append(
                                " customWidth=\"1\"");

                    if (r.getStyle() != null)
                        buf.append(" style=\"").append(r.getStyle().getId()).append("\"");

                    buf.append("/>");
                }

                buf.append("</cols>");
            }

            buf.append("<sheetData>\n");
            outputStream.write(buf.toString().getBytes("UTF-8"));
            isHeaderWritten = true;
        }
    }

    public static String escapeXML(String x) {
        final StringBuilder result = new StringBuilder();
        final StringCharacterIterator iterator = new StringCharacterIterator(x);
        char character = iterator.current();
        while (character != CharacterIterator.DONE) {
            if (character == '<') {
                result.append("&lt;");
            } else if (character == '>') {
                result.append("&gt;");
            } else if (character == '\"') {
                result.append("&quot;");
            } else if (character == '\'') {
                result.append("&#039;");
            } else if (character == '&') {
                result.append("&amp;");
            } else {
                //the char is not a special one
                //add it to the result as is
                result.append(character);
            }
            character = iterator.next();
        }
        return result.toString();
    }


    public WorkSheet appendRow(Row r) throws IOException {
        if (rowIndex >= writer.getRowsLimit()) {
            switch (writer.getRowLimitRules()) {
                case ERROR:
                    close();
                    throw new IOException("Rows number per sheet limit exceeded");

                case SKIP:
                    return this;

                case SPLIT:
                    WorkSheet clone = writer.createSplittedClone(this);
                    close();
                    return clone.appendRow(r);

                case IGNORE:
                default:
                    return doAppendRow(r);
            }
        } else
            return doAppendRow(r);
    }


    protected WorkSheet doAppendRow(Row r) throws IOException, UnsupportedEncodingException {
        writeHeaderIfDidnt();

        rowIndex++;
        StringBuilder cnt = new StringBuilder();

        cnt.append("<row r=\"").append(rowIndex).append("\"");

        if (r.getRowStyle() != null) {
            if (r.getRowStyle().getStyle() != null) {
                cnt.append(" s=\"").append(r.getRowStyle().getStyle().getId()).append("\"")
                        .append(" customFormat=\"1\"");
            }

            if (r.getRowStyle().getHeight() > 0) {
                cnt.append(" ht=\"").append(r.getRowStyle().getHeight()).append("\"").append(
                        " customHeight=\"1\"");
            }

            if (r.getRowStyle().isBorderTop()) {
                cnt.append(" thickTop=\"true\"");
            }

            if (r.getRowStyle().isBorderBottom()) {
                cnt.append(" thickBot=\"true\"");
            }
        }

        cnt.append(">");

        int cell = 0;
        for (Cell<?> c : r.getCells()) {
            if (c.isEmpty()) {
                cell++;
                continue;
            }

            CellStyle cellStyle = c.getStyle();

            if (c.getData() instanceof Number) {
                cnt.append("<c r=\"").append(columnNames[cell]).append(rowIndex).append("\"");
                if (cellStyle != null)
                    cnt.append(" s=\"").append(cellStyle.getId()).append("\"");
                cnt.append(">");

                cnt.append("<v>").append(convertNumberToRawString((Number) c.getData())).append("</v>");
                cnt.append("</c>");
            } else if (c.getData() instanceof Date) {
                cnt.append("<c r=\"").append(columnNames[cell]).append(rowIndex)
                        .append("\"");
                if (cellStyle != null)
                    cnt.append(" s=\"").append(cellStyle.getId()).append("\"");
                cnt.append(" t=\"inlineStr\"");
                cnt.append(">");

                cnt.append("<is><t>").append(DATE_TIME_FMT.format((Date) c.getData()))
                        .append("</t></is>");
                cnt.append("</c>");
            } else if (c.getData() instanceof Supplier) {
                cnt.append("<c r=\"").append(columnNames[cell]).append(rowIndex)
                        .append("\"");
                if (cellStyle != null)
                    cnt.append(" s=\"").append(cellStyle.getId()).append("\"");
                cnt.append(" t=\"inlineStr\"");
                cnt.append(">");

                cnt.append("<f>").append(escapeXML(((Supplier) c.getData()).get().toString()))
                        .append("</f>");
                cnt.append("</c>");
            } else {
                cnt.append("<c r=\"").append(columnNames[cell]).append(rowIndex)
                        .append("\" t=\"inlineStr\"");
                if (cellStyle != null)
                    cnt.append(" s=\"").append(cellStyle.getId()).append("\"");
                cnt.append(">");

                cnt.append("<is><t>").append(escapeXML(c.getData().toString())).append(
                        "</t></is>");
                cnt.append("</c>");
            }

            cell++;
        }

        cnt.append("</row>\n");

        outputStream.write(cnt.toString().getBytes("UTF-8"));

        return this;
    }

    @Override
    public void close() {
        if (outputStream != null) {
            try {
                closeCurrentSheet();
            } catch (Throwable e) {
                e.printStackTrace();
            } finally {
                try {
                    outputStream.close();
                } catch (Throwable ignored) {
                }
            }
        }

        outputStream = null;
    }

    private void closeCurrentSheet() throws IOException {
        outputStream.write("</sheetData>".getBytes());
        outputStream.write("</worksheet>".getBytes());
        columnRules.clear();
    }


    private static String convertNumberToRawString(Number number) {
        String tx = number == null ? "" : number.toString();
        return tx.replace(',', '.');
    }

    /**
     * @return the partNumber
     */
    public int getPartNumber() {
        return partNumber;
    }

    void setPartNumber(int nNumber) {
        partNumber = nNumber;
    }


    /**
     * @return the sourceName
     */
    public String getSourceName() {
        return sourceName;
    }


    /**
     * @param sourceName the sourceName to set
     */
    public void setSourceName(String sourceName) {
        this.sourceName = sourceName;
    }
}
