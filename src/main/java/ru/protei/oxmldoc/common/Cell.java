package ru.protei.oxmldoc.common;

import ru.protei.oxmldoc.style.CellStyle;

public class Cell<T> {
    private T data;
    private CellStyle style;

    public Cell(T data, CellStyle style) {
        this.data = data;
        this.style = style;
    }

    public boolean isEmpty() {
        return data == null;
    }

    /**
     * @return the data
     */
    public T getData() {
        return data;
    }

    /**
     * @return the cell style
     */
    public CellStyle getStyle() {
        return style;
    }

    public void setData(T data) {
        this.data = data;
    }

    public void setStyle(CellStyle style) {
        this.style = style;
    }
}
