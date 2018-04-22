package ru.protei.oxmldoc.style

import ru.protei.oxmldoc.style.enums.HorizontalAligment
import ru.protei.oxmldoc.style.enums.VerticalAligment

class Aligment (var horizontalAligment: HorizontalAligment, var wrapText: Boolean?, var verticalAligment: VerticalAligment) {
    constructor(horizontalAligment: HorizontalAligment) : this(horizontalAligment, false, VerticalAligment.top)
}