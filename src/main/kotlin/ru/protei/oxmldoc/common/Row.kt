package ru.protei.oxmldoc.common

import ru.protei.oxmldoc.style.CellStyle
import ru.protei.oxmldoc.style.RowStyle
import java.util.*
import kotlin.collections.ArrayList

class Row(val rowStyle: RowStyle?, val cells: MutableList<Cell<*>>) {
    constructor(): this(null, ArrayList<Cell<*>>())

    constructor(style: RowStyle) : this(style, ArrayList<Cell<*>>())

    fun append(cell: Cell<*>) {
        cells.add(cell)
    }

    fun append(x: String) {
        cells.add(Cell(x, null))
    }

    fun append(x: String, s: CellStyle) {
        cells.add(Cell(x, s))
    }

    fun append(x: Date): Unit {
        cells.add(Cell(x, null))
    }

    fun append(x: Date, s: CellStyle) {
        cells.add(Cell(x, s))
    }

    fun append(x: Number) {
        cells.add(Cell(x, null))
    }

    fun append(x: Number, s: CellStyle) {
        cells.add(Cell<Number>(x, s))
    }

    fun append(f: (String) -> String) {
        cells.add(Cell(f, null))
    }

    fun append(f: (String) -> String, s: CellStyle) {
        cells.add(Cell(f, s))
    }

    fun append() {
        cells.add(Cell(null, null))
    }

    fun clear() {
        cells.clear()
    }
}