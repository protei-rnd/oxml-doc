package ru.protei.oxmldoc.system;

import org.apache.log4j.Logger;

import java.io.*;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.List;

public class SmartStreamConsumer implements StreamConsumer {
    private static Logger log = Logger.getLogger(StreamConsumer.class);

    private final List<File> files = new ArrayList<>();
    private final String prefix = "excell";
    private final String suffix = ".wb.zip";

    public String getPrefix() {
        return prefix;
    }

    public String getSuffix() {
        return suffix;
    }

    @Override
    public void register(File file) {
        log.debug("register file on cosumer: " + file);
        files.add(file);
    }

    @Override
    public int getFilesNumber() {
        return files.size();
    }

    @Override
    public InputStream createInput() throws IOException {

        if (files.size() == 0)
            return null;

        if (files.size() == 1) {
            return new FileInputStream(files.get(0));
        }

        return new SequenceInputStream(new Enumeration<InputStream>() {
            final Iterator<File> iter = files.iterator();

            @Override
            public boolean hasMoreElements() {
                return iter.hasNext();
            }

            @Override
            public InputStream nextElement() {
                try {
                    return new FileInputStream(iter.next());
                } catch (Throwable e) {
                    return null;
                }
            }
        });
    }

    @Override
    public void release() {
        for (File f : files) {
            f.delete();
        }
        files.clear();
    }
}
