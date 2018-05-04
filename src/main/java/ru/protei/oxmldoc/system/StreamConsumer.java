package ru.protei.oxmldoc.system;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

public interface StreamConsumer {
    InputStream createInput() throws IOException;
    void register(File f);
    void release();
    String getPrefix();
    String getSuffix();
}
