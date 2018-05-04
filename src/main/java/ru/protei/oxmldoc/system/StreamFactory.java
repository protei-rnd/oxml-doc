package ru.protei.oxmldoc.system;

import java.io.IOException;
import java.io.OutputStream;

public interface StreamFactory {
    int getOpenedStreamsNumber();

    OutputStream allocate(StreamConsumer consumer) throws IOException;

    int getMaxOpenedFiles();
}
