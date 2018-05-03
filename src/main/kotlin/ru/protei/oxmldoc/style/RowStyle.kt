package ru.protei.oxmldoc.style

class RowStyle(
        val height: Int,
        val cellStyle: CellStyle?,
        val borderBottom: Boolean,
        val borderTop: Boolean
) {
    /**
     * @param height
     * @param borderBottom
     * @param borderTop
     */
    constructor(height: Int, borderBottom: Boolean, borderTop: Boolean) : this(height, null, borderBottom, borderTop)

    /**
     * @param height
     * @param style
     */
    constructor(height: Int, style: CellStyle) : this(height, style, false, false)

    /**
     * @param height
     */
    constructor(height: Int) : this(height, null, false, false)
}