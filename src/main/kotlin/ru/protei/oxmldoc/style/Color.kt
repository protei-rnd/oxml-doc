package ru.protei.oxmldoc.style

import java.awt.Color

class Color(val rgb: Int) {
    constructor(awtColor: Color): this(awtColor.rgb)
}