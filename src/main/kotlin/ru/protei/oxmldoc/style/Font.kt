package ru.protei.oxmldoc.style

class Font(
    /**
     * @return the font Id
     */
    val fontId: Int,
    val bold: Boolean,
    val colorTheme: String?,
    val color: String?,
    val italic: Boolean,
    val name: String,
    val outline: Boolean,
    val shadow: Boolean,
    val strike: Boolean,
    val underline: Boolean,
    val family: Int?,
    val charset: String?,
    val scheme: String?,
    val fontSize: String?
)