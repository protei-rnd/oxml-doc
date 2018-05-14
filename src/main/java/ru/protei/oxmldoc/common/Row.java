package ru.protei.oxmldoc.common;

import ru.protei.oxmldoc.style.CellStyle;
import ru.protei.oxmldoc.style.RowStyle;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.function.Supplier;

public class Row {
    private final List<Cell<?>> cells = new ArrayList<>();
    private RowStyle rowStyle;

    public Row(RowStyle style) {
        this.rowStyle = style;
    }

    public Row() {
    }

    public List<Cell<?>> getCells() {
        return cells;
    }

    public RowStyle getRowStyle() {
        return rowStyle;
    }

    public void setRowStyle(RowStyle rowStyle) {
        this.rowStyle = rowStyle;
    }

    public Row append(Cell<?> cell) {
        cells.add(cell);
        return this;
    }

    public Row append(String x) {
        cells.add(new Cell<>(x, null));
        return this;
    }

    public Row append(String x, CellStyle s) {
        cells.add(new Cell<>(x, s));
        return this;
    }

    public Row append(Date x) {
        cells.add(new Cell<>(x, null));
        return this;
    }

    public Row append(Date x, CellStyle s) {
        cells.add(new Cell<>(x, s));
        return this;
    }

    public Row append(Number x) {
        cells.add(new Cell<>(x, null));
        return this;
    }

    public Row append(Number x, CellStyle s) {
        cells.add(new Cell<>(x, s));
        return this;
    }

    public Row append(Supplier<String> func) {
        cells.add(new Cell<>(func, null));
        return this;
    }

    public Row append(Supplier<String> func, CellStyle s) {
        cells.add(new Cell<>(func, s));
        return this;
    }

    public Row append() {
        cells.add(new Cell<Void>(null, null));
        return this;
    }

    public void clear() {
        cells.clear();
    }
}
