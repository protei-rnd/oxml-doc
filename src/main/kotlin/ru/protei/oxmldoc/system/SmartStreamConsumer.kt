package ru.protei.oxmldoc.system

import org.apache.log4j.Logger
import java.io.*
import java.util.*

class SmartStreamConsumer: StreamConsumer {

    internal var files: MutableList<File>? = null

    override val prefix = "excell"
    override val suffix = ".wb.zip"

    init {
        files = ArrayList()
    }

    override fun register(f: File) {
        log.debug("register file on cosumer: $f")
        files!!.add(f)
    }

    fun filesNumber(): Int {
        return if (files == null) 0 else files!!.size
    }

    @Throws(IOException::class)
    fun createInput(): InputStream? {

        if (files == null || files!!.size == 0)
            return null

        return if (files!!.size == 1) {
            FileInputStream(files!![0])
        } else SequenceInputStream(object : Enumeration<InputStream> {
            internal val iter = files!!.iterator()

            override fun hasMoreElements(): Boolean {
                return iter.hasNext()
            }

            override fun nextElement(): InputStream? {
                try {
                    return FileInputStream(iter.next())
                } catch (e: Throwable) {
                    return null
                }

            }
        })

    }

    fun release() {
        for (f in files!!) {
            f.delete()
        }
        files!!.clear()
        files = null
    }

    companion object {

        private val log = Logger.getLogger(StreamConsumer::class.java)
    }
}