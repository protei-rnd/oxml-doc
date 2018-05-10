package system;

import org.junit.Assert;
import org.junit.Test;
import ru.protei.oxmldoc.system.SmartStreamConsumer;
import ru.protei.oxmldoc.system.SmartStreamFactory;
import ru.protei.oxmldoc.system.StreamConsumer;
import ru.protei.oxmldoc.system.StreamFactory;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class SmartStreamsTest {
    private class Entry {
        private StreamConsumer consumer;
        private OutputStream stream;
    }

    @Test
    public void testFactoryLimits() {
        final int TEST_NUMBER = 10;
        final int MAX_OPENED_FILES = 2;

        try {
            final StreamFactory factory = new SmartStreamFactory(MAX_OPENED_FILES);

            final List<Entry> items = new ArrayList<>();

            for (int i = 0; i < TEST_NUMBER; i++) {
                Entry e = new Entry();
                e.consumer = new SmartStreamConsumer();
                e.stream = factory.allocate(e.consumer);
                items.add(e);
            }

            items.forEach(e -> Assert.assertEquals(0, e.consumer.getFilesNumber()));

            final String testContent = "Hello World!";

            for (Entry e : items) {
                e.stream.write(("Hello").getBytes());
            }

            for (Entry e : items) {
                e.stream.write((" World!").getBytes());
            }

            for (Entry e : items)
                Assert.assertEquals(1, e.consumer.getFilesNumber());

            Assert.assertEquals(factory.getOpenedStreamsNumber(), MAX_OPENED_FILES);

            for (Entry e : items) {
                e.stream.close();
                Assert.assertEquals(1, e.consumer.getFilesNumber());

                BufferedReader reader = new BufferedReader(new InputStreamReader(e.consumer.createInput()));
                String readContent = reader.readLine();
                System.out.println("content from file: " + readContent);
                Assert.assertEquals(testContent, readContent);
                reader.close();

                e.consumer.release();
                Assert.assertEquals(e.consumer.getFilesNumber(), 0);
            }

            Assert.assertEquals(factory.getOpenedStreamsNumber(), 0);
        } catch (Throwable e) {
            e.printStackTrace();
        }
    }
}
