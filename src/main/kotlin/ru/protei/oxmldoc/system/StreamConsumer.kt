package ru.protei.oxmldoc.system

import java.io.File

interface StreamConsumer {
    fun register(f: File)

    val prefix: String

    val suffix: String
}