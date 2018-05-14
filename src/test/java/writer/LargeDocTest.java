package writer;

import org.junit.Assert;
import org.junit.Test;
import ru.protei.oxmldoc.common.Row;
import ru.protei.oxmldoc.common.RowLimitRules;
import ru.protei.oxmldoc.writer.WorkBookWriter;
import ru.protei.oxmldoc.writer.WorkSheet;

import java.io.File;

public class LargeDocTest {
    @Test
    public void writeFile() {
        try (WorkBookWriter writer = new WorkBookWriter(1000, RowLimitRules.SPLIT)) {
            WorkSheet sheet = writer.createNewSheet("sheet");
            Row row = new Row();
            for (int i = 0; i < 10000; i++) {
                for (int j = 0; j < 55; j++)
                    row.append("aaaaa");
                sheet = sheet.appendRow(row);
                row.clear();
            }
            final File bigDoc = new File("Test.xlsx");
            writer.writeTo(bigDoc);
            Assert.assertTrue(bigDoc.length() != 0);
            bigDoc.delete();
        } catch (Throwable e) {
            e.printStackTrace();
        }
    }
}
