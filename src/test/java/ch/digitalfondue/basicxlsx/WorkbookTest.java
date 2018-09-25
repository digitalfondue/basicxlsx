package ch.digitalfondue.basicxlsx;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Assert;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;

public class WorkbookTest {

    @Test
    public void testWorkbookCreation() throws IOException {

        Workbook w = new Workbook();

        Style bold = w.defineStyle().font().bold(true).build();
        Style italic = w.defineStyle().font().italic(true).build();
        Style timesNewRomanBoldAndItalic = w.defineStyle().font().name("Times New Roman").size(15).italic(true).bold(true).build();

        Sheet s = w.sheet("test");
        s.setValueAt("Hello éé èè Michał", 0, 0).withStyle(italic); //A1
        s.setValueAt("B1", 0, 1).withStyle(italic);

        s.setValueAt("A2", 1, 0).withStyle(bold);
        s.setValueAt("World!", 1, 1).withStyle(bold); //B2

        //numbers
        s.setValueAt("Numeric values", 0, 2).withStyle(timesNewRomanBoldAndItalic); //C1
        s.setValueAt(42, 1, 2); //C2
        s.setValueAt(new BigDecimal("2.512351234324832"), 2, 2); //C3
        s.setValueAt(3.14, 3, 2); //C4

        //boolean
        s.setValueAt("Boolean values", 0, 3); //D1
        s.setValueAt(true, 1, 3); //D2
        s.setValueAt(false, 2, 3); //D3

        Sheet s2 = w.sheet("test2");
        s2.setValueAt("Hello", 1, 0); //A2
        s2.setValueAt("World", 0, 1); //B1


        ByteArrayInputStream bis;
        //try (FileOutputStream fos = new FileOutputStream("test.xlsx")) {
        try (ByteArrayOutputStream os = new ByteArrayOutputStream();) {
            w.write(os);
            bis = new ByteArrayInputStream(os.toByteArray());
        }


        //check content
        org.apache.poi.ss.usermodel.Workbook workbook = WorkbookFactory.create(bis);
        Assert.assertEquals("test", workbook.getSheetName(0));
        Assert.assertEquals("test2", workbook.getSheetName(1));

        //check sheet 1 content
        org.apache.poi.ss.usermodel.Sheet sheet1 = workbook.getSheet("test");
        Assert.assertEquals("Hello éé èè Michał", sheet1.getRow(0).getCell(0).getStringCellValue());
        Assert.assertEquals(CellType.STRING, sheet1.getRow(0).getCell(0).getCellType());

        //italic, size 10, Arial (size + name = default)
        Font italicFontPoi = workbook.getFontAt(sheet1.getRow(0).getCell(0).getCellStyle().getFontIndexAsInt());
        Assert.assertTrue(italicFontPoi.getItalic());
        Assert.assertFalse(italicFontPoi.getBold());
        Assert.assertEquals(10, italicFontPoi.getFontHeightInPoints());
        Assert.assertEquals("Arial", italicFontPoi.getFontName());

        //TODO: complete check here



        //check sheet 2 content
        org.apache.poi.ss.usermodel.Sheet sheet2 = workbook.getSheet("test2");
        Assert.assertEquals("Hello", sheet2.getRow(1).getCell(0).getStringCellValue());
        Assert.assertEquals("World", sheet2.getRow(0).getCell(1).getStringCellValue());

    }
}
