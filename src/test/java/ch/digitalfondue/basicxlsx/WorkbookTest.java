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
import org.junit.Assert;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.util.Date;
import java.util.Optional;

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

        Style dateFormat = w.defineStyle().numericFormat("dd-mm-yyyy HH:mm:ss").build();

        Style borderDiagonal = w.defineStyle().diagonalColor(Style.Color.RED).diagonalStyle(Style.DiagonalStyle.BOTH)
                .border().color(Style.Color.LIME).style(Style.LineStyle.MEDIUM)
                .borderColor(Style.BorderBuilder.Border.BOTTOM, Style.Color.BLUE)
                .borderStyle(Style.BorderBuilder.Border.BOTTOM, Style.LineStyle.DASH_DOT_DOT)
                .build();

        Style rotation = w.defineStyle().rotation(90).build();

        Sheet s = w.sheet("test");
        s.setValueAt("Hello éé èè Michał", 0, 0).withStyle(italic); //A1
        s.setValueAt("B1", 0, 1).withStyle(italic); //B1

        s.setValueAt("A2", 1, 0).withStyle(bold);// A2
        s.setValueAt("World!", 1, 1).withStyle(bold); //B2

        //numbers
        s.setValueAt("Numeric values", 0, 2).withStyle(timesNewRomanBoldAndItalic); //C1
        s.setValueAt(42, 1, 2).withStyle(rotation); //C2
        s.setValueAt(new BigDecimal("2.512351234324832"), 2, 2).withStyle(twoDecimal); //C3
        s.setValueAt(3.14, 3, 2); //C4
        s.setValueAt(new BigDecimal("1.234567890"), 4, 2).withStyle(twoDecimalBuiltin); //C5

        //boolean
        s.setValueAt("Boolean values", 0, 3); //D1
        s.setValueAt(true, 1, 3); //D2
        s.setValueAt(false, 2, 3); //D3


        //
        Assert.assertTrue(s.getCellAt(0, 0).isPresent());
        Assert.assertFalse(s.getCellAt(10, 10).isPresent());

        Optional<Cell> removedCell = s.removeCellAt(0, 0);
        Assert.assertTrue(removedCell.isPresent());
        Assert.assertFalse(s.removeCellAt(0, 0).isPresent());
        s.setValueAt("Hello éé èè Michał", 0, 0).withStyle(italic); //A1
        //

        // trigger autoresize
        s.autoResizeAllColumns();
        //

        Sheet s2 = w.sheet("test2");
        s2.setValueAt("Hello", 1, 0).withStyle(borderDiagonal); //A2
        s2.setValueAt("World", 0, 1); //B1

        s2.setValueAt("Sum", 0, 2); //C1
        s2.setValueAt(1, 1, 2); //C2
        s2.setValueAt(2, 2, 2); //C3
        s2.setFormulaAt("SUM(C2:C3)", 3, 2);//C4
        s2.setFormulaAt("C2+C3+C4", 4, 2);//C5




        Date dateForSheet2Row1Col3 = DateUtils.parseDateStrictly("2015-03-07 13:26:24", "yyyy-MM-dd HH:mm:ss");
        s2.setValueAt(dateForSheet2Row1Col3, 1, 3).withStyle(dateFormat);//


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
        PoiCrossCheck.checkResult(dateForSheet2Row1Col3, bis);


    }
}
