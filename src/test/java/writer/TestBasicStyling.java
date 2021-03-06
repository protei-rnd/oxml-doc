package writer;

import org.junit.Assert;
import org.junit.Test;
import ru.protei.oxmldoc.common.ColumnRule;
import ru.protei.oxmldoc.common.Row;
import ru.protei.oxmldoc.style.*;
import ru.protei.oxmldoc.style.enums.FillPattern;
import ru.protei.oxmldoc.style.enums.HorizontalAligment;
import ru.protei.oxmldoc.style.enums.VerticalAligment;
import ru.protei.oxmldoc.writer.WorkBookWriter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class TestBasicStyling {
    @Test
    public void testBaseStyles() {
        try (WorkBookWriter writer = new WorkBookWriter()) {
            // base text font
            final Font textFont = writer.createFont();
            textFont.setFontSize("11");
            textFont.setName("Calibri");
            textFont.setColor("FF2020FF");

            // headers text font
            final Font headerFont = writer.createFont();
            headerFont.setFontSize("11");
            headerFont.setName("Calibri");
            headerFont.setColor("FF2020F0");
            headerFont.setBold(true);

            // special font for the best sold items
            final Font bestSoldFont = writer.createFont();
            bestSoldFont.setFontSize("11");
            bestSoldFont.setName("Calibri");
            bestSoldFont.setColor("FFF02020");

            // headers fill
            final Fill emptyFill = writer.createFill();

            final Fill headerFill = writer.createFill();
            headerFill.setPattern(FillPattern.SOLID);
            headerFill.setBackgroundColor(new Color(0xFFD0D0D0));
            headerFill.setFrontColor(new Color(0xFFD0D0D0));

            // cpecial fill for price column
            final Fill priceFill = writer.createFill();
            priceFill.setPattern(FillPattern.SOLID);
            priceFill.setBackgroundColor(new Color(0xFFFFFF00));
            priceFill.setFrontColor(new Color(0xFFFFFF00));


            // base style
            final CellStyle baseStyle = writer.createCellStyle();
            baseStyle.setFont(textFont);

            // header style
            final CellStyle headerStyle = writer.createCellStyle();
            headerStyle.setFont(headerFont);
            headerStyle.setFill(headerFill);
            headerStyle.setAligment(new Aligment(HorizontalAligment.CENTER, false, VerticalAligment.CENTER));

            // header row style
            final RowStyle headerRowStyle = new RowStyle(30, headerStyle, true, true);

            // price column style
            final CellStyle priceColumnStyle = writer.createCellStyle();
            priceColumnStyle.setFont(headerFont);
            priceColumnStyle.setFill(priceFill);
            priceColumnStyle.setNumberFormat(writer.createNumFormat().asNumberFmt(2));

            final CellStyle bestSoldStyle = writer.createCellStyle();
            bestSoldStyle.setFont(bestSoldFont);
            bestSoldStyle.setFill(emptyFill);

            final RowStyle bestSoldRowStyle = new RowStyle(20, bestSoldStyle, true, true);


            createWorkSheet("Market Place", writer, headerStyle, headerRowStyle, priceColumnStyle, bestSoldStyle, bestSoldRowStyle);

            final File testFile = new File("test_base.xlsx");
            try (FileOutputStream fout = new FileOutputStream(testFile)) {
                writer.writeTo(fout);
            } catch (Throwable e) {
                e.printStackTrace();
            }
            Assert.assertNotEquals(testFile.length(), 0);

            //cleaning up
            testFile.delete();
        } catch (Throwable e) {
            e.printStackTrace();
        }

    }

    private static void createWorkSheet(String name, WorkBookWriter writer, CellStyle headerStyle, RowStyle headerRowStyle, CellStyle priceColumnStyle,
                                        CellStyle bestSoldStyle, RowStyle bestSoldRowStyle) throws IOException {
        writer.createNewSheet(name);

        writer.addColumnRule(new ColumnRule(0, 50));
        writer.addColumnRule(new ColumnRule(1, 25));
        writer.addColumnRule(new ColumnRule(2, 40));
        writer.addColumnRule(new ColumnRule(3, 25, priceColumnStyle));
        writer.addColumnRule(new ColumnRule(4, 20, priceColumnStyle));

        writer.appendRow(new Row(headerRowStyle).append("Item", headerStyle).append("Count", headerStyle)
                .append("Last changes", headerStyle)
                .append("Price", headerStyle));

        writer.appendRow(new Row().append("Oranges").append("2 kg").append(new Date()).append(10).append("10.0"));
        writer.appendRow(new Row().append("Apples").append("1 kg").append(new Date()).append(20).append("20.0"));
        writer.appendRow(new Row().append("Mangos").append("50 kg.").append(new Date()).append(30).append("30.0"));

        writer.appendRow(new Row(bestSoldRowStyle).append("Bananas", bestSoldStyle)
                .append("sold out", bestSoldStyle)
                .append(new Date(), bestSoldStyle)
                .append(25.3).append("25.3"));
    }
}