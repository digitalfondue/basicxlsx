/**
 * Copyright © 2018 digitalfondue (info@digitalfondue.ch)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package ch.digitalfondue.basicxlsx;

import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.junit.Assert;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.util.Date;

public class WorkbookTest {

    private static boolean OUTPUT_FILE = false;

    @Test
    public void testWorkbookCreation() throws IOException, ParseException {

        Workbook w = new Workbook();

        Style bold = w.defineStyle().font().color("#ffcc00").bold(true).build();
        Style italic = w.defineStyle().bgColor("#ffcc00").font().color(Style.Color.GREEN).underline(Style.FontUnderlineStyle.DOUBLE_ACCOUNTING_UNDERLINE).italic(true).build();
        Style timesNewRomanBoldAndItalic = w.defineStyle().font().name("Times New Roman").size(15).italic(true).bold(true).strikeOut(true).build();

        Style twoDecimal = w.defineStyle().fgColor(Style.Color.RED).numericFormat("0.00").build();
        Style twoDecimalBuiltin = w.defineStyle().numericFormat(2).build();

        Sheet s = w.sheet("test");
        s.setValueAt("Hello éé èè Michał", 0, 0).withStyle(italic); //A1
        s.setValueAt("B1", 0, 1).withStyle(italic);

        s.setValueAt("A2", 1, 0).withStyle(bold);
        s.setValueAt("World!", 1, 1).withStyle(bold); //B2

        //numbers
        s.setValueAt("Numeric values", 0, 2).withStyle(timesNewRomanBoldAndItalic); //C1
        s.setValueAt(42, 1, 2); //C2
        s.setValueAt(new BigDecimal("2.512351234324832"), 2, 2).withStyle(twoDecimal); //C3
        s.setValueAt(3.14, 3, 2); //C4
        s.setValueAt(new BigDecimal("1.234567890"), 4, 2).withStyle(twoDecimalBuiltin); //C5

        //boolean
        s.setValueAt("Boolean values", 0, 3); //D1
        s.setValueAt(true, 1, 3); //D2
        s.setValueAt(false, 2, 3); //D3

        Sheet s2 = w.sheet("test2");
        s2.setValueAt("Hello", 1, 0); //A2
        s2.setValueAt("World", 0, 1); //B1


        Style dateFormat = w.defineStyle().numericFormat("dd-mm-yyyy HH:mm:ss").build();

        Date dateForSheet2Row1Col3 = DateUtils.parseDateStrictly("2015-03-07 13:26:24", "yyyy-MM-dd HH:mm:ss");
        s2.setValueAt(dateForSheet2Row1Col3, 1, 3).withStyle(dateFormat);


        ByteArrayInputStream bis;
        try (ByteArrayOutputStream os = new ByteArrayOutputStream();) {
            w.write(os);
            bis = new ByteArrayInputStream(os.toByteArray());

            if (OUTPUT_FILE) {
                try (FileOutputStream fos = new FileOutputStream("test.xlsx")) {
                    fos.write(os.toByteArray());
                }
            }

        }


        //check content
        org.apache.poi.ss.usermodel.Workbook workbook = WorkbookFactory.create(bis);
        Assert.assertEquals("test", workbook.getSheetName(0));
        Assert.assertEquals("test2", workbook.getSheetName(1));

        //check sheet 1 content
        org.apache.poi.ss.usermodel.Sheet sheet1 = workbook.getSheet("test");
        Assert.assertEquals("Hello éé èè Michał", sheet1.getRow(0).getCell(0).getStringCellValue());
        Assert.assertEquals(CellType.STRING, sheet1.getRow(0).getCell(0).getCellType());
        Assert.assertEquals("FFFFCC00", ((XSSFColor) sheet1.getRow(0).getCell(0).getCellStyle().getFillBackgroundColorColor()).getARGBHex());

        //italic, size 10, Arial (size + name = default)
        XSSFFont italicFontPoi = (XSSFFont) workbook.getFontAt(sheet1.getRow(0).getCell(0).getCellStyle().getFontIndexAsInt());
        Assert.assertTrue(italicFontPoi.getItalic());
        Assert.assertFalse(italicFontPoi.getBold());
        //0x008000 -> green, quite surprisingly, using getColor() return 0, but going through the XSSFColor we have the correct value, why?
        Assert.assertEquals("FF008000", italicFontPoi.getXSSFColor().getARGBHex());
        //
        Assert.assertEquals(11, italicFontPoi.getFontHeightInPoints());
        Assert.assertEquals("Calibri", italicFontPoi.getFontName());
        Assert.assertEquals(FontUnderline.DOUBLE_ACCOUNTING.getByteValue(), italicFontPoi.getUnderline());


        // Check numeric section
        // Check header
        Assert.assertEquals("Numeric values", sheet1.getRow(0).getCell(2).getStringCellValue());
        Font timesNewRomanBoldAndItalicPoi = workbook.getFontAt(sheet1.getRow(0).getCell(2).getCellStyle().getFontIndexAsInt());
        Assert.assertTrue(timesNewRomanBoldAndItalicPoi.getItalic());
        Assert.assertTrue(timesNewRomanBoldAndItalicPoi.getBold());
        Assert.assertEquals(15, timesNewRomanBoldAndItalicPoi.getFontHeightInPoints());
        Assert.assertEquals("Times New Roman", timesNewRomanBoldAndItalicPoi.getFontName());
        Assert.assertTrue(timesNewRomanBoldAndItalicPoi.getStrikeout());


        //check data format
        Assert.assertEquals(165, sheet1.getRow(2).getCell(2).getCellStyle().getDataFormat());
        Assert.assertEquals("0.00", sheet1.getRow(2).getCell(2).getCellStyle().getDataFormatString());
        Assert.assertEquals("FFFF0000", ((XSSFColor) (sheet1.getRow(2).getCell(2).getCellStyle().getFillForegroundColorColor())).getARGBHex());
        //
        Assert.assertEquals(2, sheet1.getRow(4).getCell(2).getCellStyle().getDataFormat());
        Assert.assertEquals("0.00", sheet1.getRow(4).getCell(2).getCellStyle().getDataFormatString());
        //


        //check sheet 2 content
        org.apache.poi.ss.usermodel.Sheet sheet2 = workbook.getSheet("test2");
        Assert.assertEquals("Hello", sheet2.getRow(1).getCell(0).getStringCellValue());
        Assert.assertEquals("World", sheet2.getRow(0).getCell(1).getStringCellValue());

        //check date
        Assert.assertEquals(dateForSheet2Row1Col3, sheet2.getRow(1).getCell(3).getDateCellValue());
        Assert.assertEquals("dd-mm-yyyy HH:mm:ss", sheet2.getRow(1).getCell(3).getCellStyle().getDataFormatString());
    }
}
