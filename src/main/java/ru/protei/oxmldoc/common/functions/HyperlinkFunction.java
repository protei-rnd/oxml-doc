package ru.protei.oxmldoc.common.functions;

import java.util.function.Supplier;

public class HyperlinkFunction implements Supplier<String> {
    private final String url;
    private final String linkName;

    /**
     * @param url
     * @param linkName
     */
    public HyperlinkFunction(String url, String linkName) {
        this.url = url;
        this.linkName = linkName;
    }

    /**
     * @param sheetName
     * @param field
     * @param linkName
     */
    public HyperlinkFunction(String sheetName, String field, String linkName) {
        this.url = "#" + sheetName + "!" + (field != null && field.trim().length() > 0 ? field : "A1");
        this.linkName = linkName;
    }

    @Override
    public String get() {
        return "HYPERLINK(\"" + url + "\"" + (linkName != null && linkName.trim().length() > 0 ? ",\"" + linkName + "\"" : "") + ")";
    }
}
