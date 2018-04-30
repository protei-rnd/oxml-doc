package ru.protei.oxmldoc.style

import ru.protei.oxmldoc.style.enums.FillPattern

data class Fill(val id: Int, val fillPattern: FillPattern, val backgroundColor: Color, val frontColor: Color)