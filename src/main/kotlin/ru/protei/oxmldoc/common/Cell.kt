package ru.protei.oxmldoc.common

import ru.protei.oxmldoc.style.CellStyle

class Cell<T>(//  private int m_Index;
        /**
         * @return the data
         */
        val data: T?,
        /**
         * @return the style
         */
        val style: CellStyle?) {

    val isEmpty: Boolean
        get() = data == null
}