package ru.protei.oxmldoc.system

import org.apache.log4j.Logger
import java.io.File
import java.io.FileOutputStream
import java.io.IOException
import java.io.OutputStream
import java.util.*
import java.util.concurrent.locks.ReentrantLock

class SmartStreamFactory(override val maxOpenedFiles: Int) : StreamFactory {

    private val lock = ReentrantLock()

    private val connInfo: Queue<ConnInfo>

    override val openedStreamsNumber: Int
        get() {
            lock.lock()
            try {
                return this.connInfo.size
            } finally {
                lock.unlock()
            }
        }

    init {
        connInfo = LinkedList()
    }

    @Throws(IOException::class)
    override fun allocate(consumer: StreamConsumer): OutputStream {
        return SmartOutputStream(consumer)
    }

    private fun factCloseStream(handler: ConnInfo) {
        lock.lock()
        try {
            if (connInfo.contains(handler)) {
                doCloseStream(handler)
                connInfo.remove(handler)
            }
        } finally {
            lock.unlock()
        }
    }

    @Throws(IOException::class)
    private fun factOpenStream(caller: SmartOutputStream, info: ConnInfo): Boolean {
        lock.lock()

        try {
            if (connInfo.size >= maxOpenedFiles) {
                val old = connInfo.poll()
                log.debug("close oldest file: " + old!!.f!!)
                doCloseStream(old)
            }

            connInfo.offer(info)

            caller.attachStream(FileOutputStream(info.f!!, true))
            return true
        } finally {
            lock.unlock()
        }
    }

    private fun doCloseStream(info: ConnInfo) {
        info.stream.detachStream()
        try {
            info.out!!.flush()
        } catch (e: Throwable) {
        }

        try {
            info.out!!.close()
        } catch (e: Throwable) {
        }

    }


    internal inner class ConnInfo(var stream: SmartOutputStream) {
        var allocated: Long = 0
        var out: OutputStream? = null
        var f: File? = null

        init {
            this.allocated = System.currentTimeMillis()
        }

    }

    inner class SmartOutputStream @Throws(IOException::class)
    internal constructor(internal var consumer: StreamConsumer) : OutputStream() {
        internal var handler: ConnInfo? = null

        private val output: OutputStream?
            @Throws(IOException::class)
            get() {
                if (this.handler!!.f == null) {
                    this.handler!!.f = File.createTempFile(consumer.prefix, consumer.suffix)

                    log.debug("create temp file " + this.handler!!.f!!)

                    consumer.register(this.handler!!.f!!)
                }

                if (this.handler!!.out == null && !factOpenStream(this, this.handler!!))
                    throw IOException("unable to open stream")
                return handler!!.out
            }

        init {
            this.handler = ConnInfo(this)
        }

        @Synchronized
        internal fun attachStream(out: OutputStream) {
            this.handler!!.out = out
        }

        @Synchronized
        internal fun detachStream() {
            this.handler!!.out = null
        }

        @Synchronized
        @Throws(IOException::class)
        override fun write(b: Int) {
            output!!.write(b)
        }

        @Synchronized
        @Throws(IOException::class)
        override fun write(b: ByteArray) {
            output!!.write(b)
        }

        @Synchronized
        @Throws(IOException::class)
        override fun write(b: ByteArray, off: Int, len: Int) {
            output!!.write(b, off, len)
        }

        @Synchronized
        @Throws(IOException::class)
        override fun flush() {
            output!!.flush()
        }

        @Synchronized
        @Throws(IOException::class)
        override fun close() {
            if (this.handler != null) {
                factCloseStream(this.handler!!)
                this.handler = null
            }
        }
    }

    companion object {

        private val log = Logger.getLogger(SmartStreamFactory::class.java)

        val DEFAULT: SmartStreamFactory = SmartStreamFactory(50)
    }
}