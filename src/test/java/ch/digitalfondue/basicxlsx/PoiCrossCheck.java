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

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.junit.Assert;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.Date;

import static org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle.INT_THIN;

public class PoiCrossCheck {

    public static void checkResult(Date dateForSheet2Row1Col3, ByteArrayInputStream bis) throws IOException {
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

        // check rotation
        Assert.assertEquals(85, sheet1.getRow(1).getCell(2).getCellStyle().getRotation());
        //


        //check sheet 2 content
        org.apache.poi.ss.usermodel.Sheet sheet2 = workbook.getSheet("test2");
        Assert.assertEquals("Hello", sheet2.getRow(1).getCell(0).getStringCellValue());

        //http://apache-poi.1045710.n5.nabble.com/Diagonal-border-td5720338.html
        XSSFCellStyle style = (XSSFCellStyle) workbook.getCellStyleAt(sheet2.getRow(1).getCell(0).getCellStyle().getIndex());
        StylesTable stylesTable = ((XSSFWorkbook) workbook).getStylesSource();
        XSSFCellBorder border = stylesTable.getBorderAt((int) style.getCoreXf().getBorderId());
        Assert.assertEquals(true, border.getCTBorder().getDiagonalDown());
        Assert.assertEquals(true, border.getCTBorder().getDiagonalUp());
        Assert.assertEquals(STBorderStyle.Enum.forString("thin"), border.getCTBorder().getDiagonal().getStyle());
        Assert.assertEquals("<xml-fragment rgb=\"FFFF0000\"/>", border.getCTBorder().getDiagonal().getColor().toString());

        //should be for all the borders

        Assert.assertEquals(STBorderStyle.Enum.forString("medium"), border.getCTBorder().getTop().getStyle());
        Assert.assertEquals("<xml-fragment rgb=\"FF00FF00\"/>", border.getCTBorder().getTop().getColor().toString());

        Assert.assertEquals(STBorderStyle.Enum.forString("medium"), border.getCTBorder().getRight().getStyle());
        Assert.assertEquals("<xml-fragment rgb=\"FF00FF00\"/>", border.getCTBorder().getRight().getColor().toString());

        Assert.assertEquals(STBorderStyle.Enum.forString("dashDotDot"), border.getCTBorder().getBottom().getStyle());
        Assert.assertEquals("<xml-fragment rgb=\"FF0000FF\"/>", border.getCTBorder().getBottom().getColor().toString()); //blue

        Assert.assertEquals(STBorderStyle.Enum.forString("medium"), border.getCTBorder().getLeft().getStyle());
        Assert.assertEquals("<xml-fragment rgb=\"FF00FF00\"/>", border.getCTBorder().getLeft().getColor().toString());

        //
        Assert.assertEquals("World", sheet2.getRow(0).getCell(1).getStringCellValue());

        //check date
        Assert.assertEquals(dateForSheet2Row1Col3, sheet2.getRow(1).getCell(3).getDateCellValue());
        Assert.assertEquals("dd-mm-yyyy HH:mm:ss", sheet2.getRow(1).getCell(3).getCellStyle().getDataFormatString());
    }
}
