package ru.protei.oxmldoc.system;

import org.apache.log4j.Logger;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.LinkedList;
import java.util.Queue;
import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReentrantLock;

public class SmartStreamFactory implements StreamFactory {
    private static Logger log = Logger.getLogger(SmartStreamFactory.class);

    public static final SmartStreamFactory DEFAULT = new SmartStreamFactory(50);

    private Lock lock = new ReentrantLock();

    private int maxOpenedFiles;

    private Queue<ConnInfo> connInfo;

    public SmartStreamFactory(int maxOpenedFiles) {
        this.maxOpenedFiles = maxOpenedFiles;
        connInfo = new LinkedList<>();
    }

    public int getMaxOpenedFiles() {
        return maxOpenedFiles;
    }

    @Override
    public int getOpenedStreamsNumber() {
        lock.lock();
        try {
            return this.connInfo.size();
        } finally {
            lock.unlock();
        }
    }

    @Override
    public OutputStream allocate(StreamConsumer consumer) throws IOException {
        return new SmartOutputStream(consumer);
    }

    private void factCloseStream(ConnInfo handler) {
        lock.lock();
        try {
            if (connInfo.contains(handler)) {
                doCloseStream(handler);
                connInfo.remove(handler);
            }
        } finally {
            lock.unlock();
        }
    }

    private boolean factOpenStream(SmartOutputStream caller, ConnInfo info) throws IOException {
        lock.lock();

        try {
            if (connInfo.size() >= maxOpenedFiles) {
                ConnInfo old = connInfo.poll();
                log.debug("close oldest file: " + old.f);
                doCloseStream(old);
            }

            connInfo.offer(info);

            caller.attachStream(new FileOutputStream(info.f, true));
            return true;
        } finally {
            lock.unlock();
        }
    }

    private void doCloseStream(ConnInfo info) {
        info.stream.detachStream();
        try {
            info.out.flush();
        } catch (Throwable e) {
        }
        try {
            info.out.close();
        } catch (Throwable e) {
        }
    }


    class ConnInfo {
        long allocated;
        OutputStream out;
        SmartOutputStream stream;
        File f;

        public ConnInfo(SmartOutputStream stream) {
            this.allocated = System.currentTimeMillis();
            this.stream = stream;
        }

    }

    private class SmartOutputStream extends OutputStream {
        StreamConsumer consumer;
        ConnInfo handler;

        SmartOutputStream(StreamConsumer consumer) throws IOException {
            this.consumer = consumer;
            this.handler = new ConnInfo(this);
        }

        private OutputStream getOutput() throws IOException {
            if (this.handler.f == null) {
                this.handler.f = File.createTempFile(consumer.getPrefix(), consumer.getSuffix());

                log.debug("create temp file " + this.handler.f);

                consumer.register(this.handler.f);
            }

            if (this.handler.out == null && !factOpenStream(this, this.handler))
                throw new IOException("unable to open stream");
            return handler.out;
        }

        synchronized void attachStream(OutputStream out) {
            this.handler.out = out;
        }

        synchronized void detachStream() {
            this.handler.out = null;
        }

        @Override
        public synchronized void write(int b) throws IOException {
            getOutput().write(b);
        }

        @Override
        public synchronized void write(byte[] b) throws IOException {
            getOutput().write(b);
        }

        @Override
        public synchronized void write(byte[] b, int off, int len) throws IOException {
            getOutput().write(b, off, len);
        }

        @Override
        public synchronized void flush() throws IOException {
            getOutput().flush();
        }

        @Override
        public synchronized void close() throws IOException {
            if (this.handler != null) {
                factCloseStream(this.handler);
                this.handler = null;
            }
        }
    }
}
