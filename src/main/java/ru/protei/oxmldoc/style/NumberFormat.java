package ru.protei.oxmldoc.style;

public class NumberFormat {
    private final int id;
    private String format;

    public NumberFormat(int id) {
        this.id = id;
    }

    private static StringBuilder mappend(StringBuilder sb, char appendedChar, int times) {
        while (times > 0) {
            sb.append(appendedChar);
            times--;
        }
        return sb;
    }

    public int getId() {
        return id;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public NumberFormat asNumberFmt(int fractionDigits) {
        return asNumberFmt(1, fractionDigits);
    }

    private NumberFormat asNumberFmt(int integerDigits, int fractionDigits) {
        StringBuilder sb = mappend(new StringBuilder(), '0', integerDigits).append('.');
        this.format = mappend(sb, '0', fractionDigits).toString();
        return this;
    }
}
