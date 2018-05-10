package ru.protei.oxmldoc.reader;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import ru.protei.oxmldoc.reader.enums.XSSFDataType;
import ru.protei.oxmldoc.reader.enums.XSSFState;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedList;

public class XSSFListener extends DefaultHandler {
    private StylesTable stylesTable;
    private ReadOnlySharedStringsTable readOnlySharedStringsTable;

    private ExcelHandler excelHandler;

    private XSSFState xssfState;
    private CellRangeAddress cellRangeAddress;
    private LinkedList<ParserNode> parserNodes;
    private RowData currentRow;
    private CellInfo cellInfo;

    private static class CellInfo {
        XSSFDataType dataType = XSSFDataType.NUMBER;
        boolean isDate = false;
        CellReference ref;
        StringBuilder builder = new StringBuilder();
    }

    private static class RowData {
        private CellValue[] cellValues;
        private int firstColumn;

        RowData(CellRangeAddress rec) {
            int size = rec.getLastColumn() - rec.getFirstColumn() + 1;
            cellValues = new CellValue[size];
            for (int i = 0; i < size; i++)
                cellValues[i] = null;
            firstColumn = rec.getFirstColumn();
        }

        void setValue(int index, CellValue value) {
            index -= firstColumn;
            if (index < cellValues.length)
                cellValues[index] = value;
        }

        void export(ExcelHandler handler) {
            for (int i = 0; i < cellValues.length; i++)
                handler.setCell(i + firstColumn, cellValues[i]);
        }

        public String toString() {
            if ((cellValues == null) || (cellValues.length <= 0))
                return "[]";

            StringBuilder sb = new StringBuilder();
            sb.append('[');
            for (int i = 0; i < cellValues.length; i++) {
                if (i > 0)
                    sb.append(", ");
                sb.append(String.valueOf(cellValues[i]));
            }
            sb.append(']');
            return sb.toString();
        }
    }

    private static class ParserNode {
        private String name;
        private XSSFState location;

        ParserNode(String name, Attributes atts, XSSFState loc) {
            this.name = name;
            location = loc;
        }

        String getName() {
            return name;
        }

        XSSFState getLocation() {
            return location;
        }
    }

    /* (non-Javadoc)
     * @see org.xml.sax.helpers.DefaultHandler#startElement(java.lang.String, java.lang.String, java.lang.String, org.xml.sax.Attributes)
     */
    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
        // Store current path
        parserNodes.addLast(new ParserNode(name, attributes, xssfState));

        switch (xssfState) {
            case ROOT:
                if ("worksheet".equals(name))
                    xssfState = XSSFState.WORKSHEET;
                else
                    xssfState = XSSFState.SKIP;
                break;
            case WORKSHEET:
                if ("dimension".equals(name)) {
                    cellRangeAddress = CellRangeAddress.valueOf(attributes.getValue("ref"));
                    xssfState = XSSFState.SKIP;
                } else if ("sheetData".equals(name))
                    xssfState = XSSFState.SHEET_DATA;
                else
                    xssfState = XSSFState.SKIP;
                break;
            case SHEET_DATA:
                if ("row".equals(name)) // Row
                {
                    int rowid = Integer.valueOf(attributes.getValue("r")) - 1;
                    excelHandler.beginRow(rowid, cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
                    currentRow = new RowData(cellRangeAddress);
                    xssfState = XSSFState.ROW;
                } else
                    xssfState = XSSFState.SKIP;
                break;
            case ROW:
                if ("c".equals(name)) // Column
                {
                    cellInfo = new CellInfo();
                    cellInfo.dataType = XSSFDataType.parse(attributes.getValue("t"));

                    if (cellInfo.dataType != XSSFDataType.ERROR) {
                        cellInfo.ref = new CellReference(attributes.getValue("r"));
                        String style = attributes.getValue("s");

                        if (style != null) {
                            int styleIndex = Integer.valueOf(style);
                            XSSFCellStyle styleRef = stylesTable.getStyleAt(styleIndex);
                            int formatIndex = styleRef.getDataFormat();
                            String formatString = styleRef.getDataFormatString();

                            if (formatString == null)
                                formatString = BuiltinFormats.getBuiltinFormat(formatIndex);

                            cellInfo.isDate = ((formatString != null) && (DateUtil.isADateFormat(formatIndex, formatString)));
                        }

                        xssfState = XSSFState.CELL;
                    } else
                        xssfState = XSSFState.SKIP;
                } else
                    xssfState = XSSFState.SKIP;
                break;
            case CELL:
                if ("v".equals(name) || "inlineStr".equals(name)) // Value
                    xssfState = XSSFState.CELL_VALUE;
                else
                    xssfState = XSSFState.SKIP;
                break;
        }

    }

    /* (non-Javadoc)
     * @see org.xml.sax.helpers.DefaultHandler#endElement(java.lang.String, java.lang.String, java.lang.String)
     */
    public void endElement(String uri, String localName, String name) throws SAXException {
        ParserNode pnode = parserNodes.removeLast();
        if (!pnode.getName().equals(name))
            throw new SAXException("Invalid document format");

        switch (xssfState) {
            case CELL:
                int column = cellInfo.ref.getCol();
                String strvalue = cellInfo.builder.toString();

                switch (cellInfo.dataType) {
                    case BOOL:
                        currentRow.setValue(column, new CellValue(!"0".equals(strvalue), strvalue));
                        break;
                    case INLINESTR:
                    case FORMULA:
                        XSSFRichTextString rtsi = new XSSFRichTextString(strvalue);
                        currentRow.setValue(column, new CellValue(rtsi.toString()));
                        break;
                    case SSTINDEX:
                        try {
                            int idx = Integer.parseInt(strvalue);
                            XSSFRichTextString rtss = new XSSFRichTextString(readOnlySharedStringsTable.getEntryAt(idx));
                            currentRow.setValue(column, new CellValue(rtss.toString()));
                        } catch (NumberFormatException ex) {
                            throw new SAXException("Invalid document format", ex);
                        }
                        break;
                    case NUMBER:
                        CellValue cellValue = new CellValue(
                                cellInfo.isDate ? CellValue.DATE : CellValue.NUMBER,
                                strvalue, strvalue.length() > 0 ? Double.valueOf(strvalue) : null
                        );
                        currentRow.setValue(column, cellValue);
                        break;
                }
                cellInfo = null;
                break;

            case ROW:
                currentRow.export(excelHandler);
                excelHandler.endRow();
                currentRow = null;
                break;
        }

        xssfState = pnode.getLocation();
    }

    /**
     * Captures characters only if a suitable element is open.
     * Originally was just "v"; extended for inlineStr also.
     */
    public void characters(char[] ch, int start, int length) throws SAXException {
        if (xssfState == XSSFState.CELL_VALUE)
            cellInfo.builder.append(ch, start, length);
    }

    public XSSFListener(ExcelHandler handler, StylesTable styles, ReadOnlySharedStringsTable strings) {
        stylesTable = styles;
        readOnlySharedStringsTable = strings;
        excelHandler = handler;
        parserNodes = new LinkedList<>();
    }

    public void processSheet(InputStream sheetInputStream) throws IOException {
        xssfState = XSSFState.ROOT;
        cellRangeAddress = null;
        currentRow = null;
        cellInfo = null;
        parserNodes.clear();

        try {
            InputSource sheetSource = new InputSource(sheetInputStream);
            SAXParserFactory saxFactory = SAXParserFactory.newInstance();
            SAXParser saxParser;

            saxParser = saxFactory.newSAXParser();

            XMLReader sheetParser = saxParser.getXMLReader();

            sheetParser.setContentHandler(this);
            sheetParser.parse(sheetSource);
        } catch (NullPointerException | ParserConfigurationException | SAXException ex) {
            throw new IOException("Invalid document format", ex);
        }
    }

}
