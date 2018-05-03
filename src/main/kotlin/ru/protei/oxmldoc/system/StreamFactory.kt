package ru.protei.oxmldoc.system

import java.io.IOException
import java.io.OutputStream

interface StreamFactory {

    val openedStreamsNumber: Int

    val maxOpenedFiles: Int

    @Throws(IOException::class)
    fun allocate(consumer: StreamConsumer): OutputStream
}