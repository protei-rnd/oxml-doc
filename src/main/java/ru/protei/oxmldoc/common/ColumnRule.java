package ru.protei.oxmldoc.common;

import ru.protei.oxmldoc.style.CellStyle;

public class ColumnRule {
    private int begin;
    private int end;
    private int width;
    private CellStyle style;

    public ColumnRule(int col, int width) {
        this.begin = col;
        this.end = col;
        this.width = width;
    }

    public ColumnRule(int col, int width, CellStyle style) {
        this.begin = col;
        this.end = col;
        this.width = width;
        this.style = style;
    }

    /**
     * @param begin
     * @param end
     * @param width
     */
    public ColumnRule(int begin, int end, int width) {
        this.begin = begin;
        this.end = end;
        this.width = width;
    }

    /**
     * @param begin
     * @param end
     * @param width
     * @param style
     */
    public ColumnRule(int begin, int end, int width, CellStyle style) {
        this.begin = begin;
        this.end = end;
        this.width = width;
        this.style = style;
    }

    public int getBegin() {
        return begin;
    }

    public void setBegin(int begin) {
        this.begin = begin;
    }

    public int getEnd() {
        return end;
    }

    public void setEnd(int end) {
        this.end = end;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public CellStyle getStyle() {
        return style;
    }

    public void setStyle(CellStyle style) {
        this.style = style;
    }
}
