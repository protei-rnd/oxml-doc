package ru.protei.oxmldoc.reader.enums;

import java.util.Arrays;

public enum XSSFDataType {
    BOOL("b"),
    ERROR("e"),
    FORMULA("str"),
    INLINESTR("inlineStr"),
    SSTINDEX("s"),
    NUMBER("n");

    private String type;

    private XSSFDataType(String type) {
        this.type = type;
    }

    /**
     * Decode cell data type or throw exception
     *
     * @param src string presentation of data type in ooxml
     * @return decoded data type
     */
    public static XSSFDataType parse(String src) {
        return Arrays.stream(XSSFDataType.values()).filter(dt -> dt.type.equalsIgnoreCase(src)).findFirst().orElse(XSSFDataType.NUMBER);
    }
}
