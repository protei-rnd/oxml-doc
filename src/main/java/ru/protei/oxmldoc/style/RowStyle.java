package ru.protei.oxmldoc.style;

public class RowStyle {
    private int height;
    private CellStyle style;
    private boolean borderBottom;
    private boolean borderTop;

    /**
     * @param height
     * @param style
     * @param borderBottom
     * @param borderTop
     */
    public RowStyle(int height, CellStyle style, boolean borderBottom, boolean borderTop) {
        this.height = height;
        this.style = style;
        this.borderBottom = borderBottom;
        this.borderTop = borderTop;
    }

    /**
     * @param height
     * @param borderBottom
     * @param borderTop
     */
    public RowStyle(int height, boolean borderBottom, boolean borderTop) {
        this.height = height;
        this.borderBottom = borderBottom;
        this.borderTop = borderTop;
    }

    /**
     * @param height
     * @param style
     */
    public RowStyle(int height, CellStyle style) {
        this.height = height;
        this.style = style;
    }

    /**
     * @param height
     */
    public RowStyle(int height) {
        this.height = height;
    }

    public int getHeight() {
        return height;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public CellStyle getStyle() {
        return style;
    }

    public void setStyle(CellStyle style) {
        this.style = style;
    }

    public boolean isBorderBottom() {
        return borderBottom;
    }

    public void setBorderBottom(boolean borderBottom) {
        this.borderBottom = borderBottom;
    }

    public boolean isBorderTop() {
        return borderTop;
    }

    public void setBorderTop(boolean borderTop) {
        this.borderTop = borderTop;
    }
}
