package ru.protei.oxmldoc.common

import ru.protei.oxmldoc.style.CellStyle

/**
 * @param begin
 * @param end
 * @param width
 * @param style
 */
class ColumnRule(val begin: Int?, val end: Int?, val width: Int?, val style: CellStyle?) {
    constructor(col: Int, width: Int): this(col, col, width, null)

    constructor(col: Int, width: Int, style: CellStyle): this(col, col, width, style)

    /**
     * @param begin
     * @param end
     * @param width
     */
    constructor(begin: Int, end: Int, width: Int) : this(begin, end, width, null)
}