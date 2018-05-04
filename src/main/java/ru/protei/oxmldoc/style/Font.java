package ru.protei.oxmldoc.style;

public class Font {
    private final int fontId;

    public Font (int index){
        fontId = index;
    }

    private boolean isBold;
    private String colorTheme;
    private String color;
    private boolean isItalic;
    private String name;
    private boolean isOutlined;
    private boolean isShadowed;
    private boolean isStriked;
    private boolean isUnderlined;
    private int family;
    private String charset;
    private String scheme;
    private String fontSize;

    /**
     * @return the fontId
     */
    public int getFontId()
    {
        return fontId;
    }

    public boolean isBold() {
        return isBold;
    }

    public void setBold(boolean bold) {
        this.isBold = bold;
    }

    public String getColorTheme() {
        return colorTheme;
    }

    public void setColorTheme(String colorTheme) {
        this.colorTheme = colorTheme;
    }

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public boolean isItalic() {
        return isItalic;
    }

    public void setItalic(boolean italic) {
        this.isItalic = italic;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public boolean isOutlined() {
        return isOutlined;
    }

    public void setOutlined(boolean outlined) {
        this.isOutlined = outlined;
    }

    public boolean isShadowed() {
        return isShadowed;
    }

    public void setShadowed(boolean shadowed) {
        this.isShadowed = shadowed;
    }

    public boolean isStriked() {
        return isStriked;
    }

    public void setIsStriked(boolean isStriked) {
        this.isStriked = isStriked;
    }

    public boolean isUnderlined() {
        return isUnderlined;
    }

    public void setUnderlined(boolean underlined) {
        this.isUnderlined = underlined;
    }

    public int getFamily() {
        return family;
    }

    public void setFamily(int family) {
        this.family = family;
    }

    public String getCharset() {
        return charset;
    }

    public void setCharset(String charset) {
        this.charset = charset;
    }

    public String getScheme() {
        return scheme;
    }

    public void setScheme(String scheme) {
        this.scheme = scheme;
    }

    public String getFontSize() {
        return fontSize;
    }

    public void setFontSize(String fontSize) {
        this.fontSize = fontSize;
    }
}
