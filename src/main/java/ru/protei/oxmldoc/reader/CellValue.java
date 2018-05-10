package ru.protei.oxmldoc.reader;

import org.apache.poi.ss.usermodel.DateUtil;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class CellValue {
    private static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    /**
     * Cell type indicates that value stored in cell is a numeric value
     */
    static final int NUMBER = 0;

    /**
     * Cell type indicates that value stored in cell is a logical value
     */
    static final int BOOLEAN = 1;

    /**
     * Cell type indicates that value stored in cell is a character string value
     */
    static final int STRING = 2;

    /**
     * Cell type indicates that value stored in cell is a date value
     */
    static final int DATE = 3;

    private int type;
    private Object data;
    private String parsedValue;

    /**
     * Constructs logical cell value
     *
     * @param data value
     */
    public CellValue(boolean data, String parsedValue) {
        type = BOOLEAN;
        this.data = data;
        this.parsedValue = parsedValue;
    }

    /**
     * Constructs character string cell value
     *
     * @param data character string
     */
    public CellValue(String data) {
        type = STRING;
        this.data = data;
        this.parsedValue = data;
    }

    /**
     * Constructs date cell value
     *
     * @param data date value
     */
    public CellValue(Date data, String parsedValue) {
        type = DATE;
        this.data = DateUtil.getExcelDate(data);
        this.parsedValue = parsedValue;
    }

    /**
     * Constructs numeric cell value
     *
     * @param data numeric cell value
     */
    public CellValue(double data, String pv) {
        type = NUMBER;
        this.data = data;
        this.parsedValue = pv;
    }

    /**
     * Constructs cell value with specified type
     *
     * @param type cell value type
     * @param data cell value
     */
    public CellValue(int type, String parsedValue, Object data) {
        switch (type) {
            case BOOLEAN:
                if (data != null) {
                    if (!(data instanceof Boolean))
                        throw new IllegalArgumentException("Invalid data");
                }
                break;
            case NUMBER:
            case DATE:
                if (data != null) {
                    if (!(data instanceof Double))
                        throw new IllegalArgumentException("Invalid data");
                }
                break;
            case STRING:
                if (data != null) {
                    if (!(data instanceof String))
                        throw new IllegalArgumentException("Invalid data");
                }
                break;
            default:
                throw new IllegalArgumentException("Invalid cell type");
        }
        this.data = data;
        this.type = type;
        this.parsedValue = parsedValue;
    }

    public String getParsedValue() {
        return parsedValue;
    }

    /**
     * Return cell type
     *
     * @return cell type
     */
    public int getType() {
        return type;
    }

    /**
     * Check if value is null
     *
     * @return true if cell value is NULL
     */
    public boolean isNull() {
        return data == null;
    }

    /**
     * Check if cell value is numeric
     *
     * @return true if cell value if numeric
     */
    public boolean isNumber() {
        return type == NUMBER;
    }

    /**
     * Check if cell value is logical
     *
     * @return true if cell value is logical
     */
    public boolean isBoolean() {
        return type == BOOLEAN;
    }

    /**
     * Check if cell value is a string
     *
     * @return true if cell value is a string
     */
    public boolean isString() {
        return type == STRING;
    }

    /**
     * Check if cell value is a date
     *
     * @return true if cell value is a date
     */
    public boolean isDate() {
        return type == DATE;
    }

    /**
     * Get cell value as a double
     *
     * @return cell value
     * @throws NullPointerException  if cell value is null
     * @throws NumberFormatException if cell value can not be converted to double
     */
    public double getDouble() {
        if (data == null)
            throw new NullPointerException();
        switch (type) {
            case BOOLEAN:
                return ((Boolean) data) ? 1.0 : 0.0;
            case DATE:
            case NUMBER:
                return (Double) data;

            default:
                return Double.parseDouble((String) data);
        }
    }

    /**
     * Get cell value as a long
     *
     * @return cell value
     * @throws NullPointerException  if cell value is null
     * @throws NumberFormatException if cell value can not be converted to long
     */
    public long getLong() {
        return (long) getDouble();
    }

    /**
     * Get cell value as an integer
     *
     * @return cell value
     * @throws NullPointerException  if cell value is null
     * @throws NumberFormatException if cell value can not be converted to integer
     */
    public int getInteger() {
        return (int) getDouble();
    }

    /**
     * Get cell value as short
     *
     * @return cell value
     * @throws NullPointerException  if cell value is null
     * @throws NumberFormatException if cell value can not be converted to short
     */
    public short getShort() {
        return (short) getDouble();
    }

    /**
     * Get cell value as byte
     *
     * @return cell value
     * @throws NullPointerException  if cell value is null
     * @throws NumberFormatException if cell value can not be converted to byte
     */
    public short getByte() {
        return (byte) getDouble();
    }

    /**
     * Get cell value as boolean
     *
     * @return cell value
     * @throws NullPointerException if cell value is null
     */
    public boolean getBoolean() {
        if (data == null)
            throw new NullPointerException();
        switch (type) {
            case BOOLEAN:
                return (Boolean) data;
            case DATE:
            case NUMBER:
                return (Double) data != 0.0;

            default:
                return Boolean.parseBoolean((String) data);
        }
    }

    /**
     * Get cell value as date
     *
     * @return cell value, null if cell value is null or can not be parsed as date
     */
    public Date getDate() {
        if (data == null)
            return null;
        switch (type) {
            case BOOLEAN:
                return null;

            case DATE:
            case NUMBER:
                return DateUtil.getJavaDate((Double) data);

            default:
                try {
                    return sdf.parse((String) data);
                } catch (ParseException ex) {
                    return null;
                }
        }
    }

    /**
     * Get string value
     *
     * @return string value, can be null
     */
    public String getString() {
        if (data != null) {
            switch (type) {
                case DATE:
                    return String.valueOf(DateUtil.getJavaDate((Double) data));
                case NUMBER:
                case STRING:
                case BOOLEAN:
                    return String.valueOf(data);
            }
        }

        return null;
    }

    public String toString() {
        if (data != null) {
            switch (type) {
                case DATE:
                    return String.valueOf(DateUtil.getJavaDate((Double) data));
                case NUMBER:
                case STRING:
                case BOOLEAN:
                    return String.valueOf(data);
            }
        }
        return String.valueOf(null);
    }
}
