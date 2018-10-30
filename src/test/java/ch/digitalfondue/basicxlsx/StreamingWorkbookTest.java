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
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.util.Date;
import java.util.stream.Stream;

import static ch.digitalfondue.basicxlsx.Cell.cell;
import static ch.digitalfondue.basicxlsx.Cell.formula;
import static ch.digitalfondue.basicxlsx.StreamingWorkbook.row;

public class StreamingWorkbookTest {
    private static boolean OUTPUT_FILE = false;

    @Test
    public void testWorkbookCreation() throws IOException, ParseException {

        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        Date dateForSheet2Row1Col3 = DateUtils.parseDateStrictly("2015-03-07 13:26:24", "yyyy-MM-dd HH:mm:ss");

        try (StreamingWorkbook w = new StreamingWorkbook(baos)) {
            // you must define the styles before
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
            //

            //sheet 1
            Cell[] s1row1 = new Cell[] {
                    cell("Hello éé èè Michał").withStyle(italic), //A1
                    cell("B1").withStyle(italic), //B1
                    cell("Numeric values").withStyle(timesNewRomanBoldAndItalic), //C1
                    cell("Boolean values") //D1
            };

            Cell[] s1row2 = new Cell[]{
                    cell("A2").withStyle(bold), //A2
                    cell("World!").withStyle(bold), //B2
                    cell(42), //C2
                    cell(true) //D2
            };

            Cell[] s1row3 = new Cell[] {
                    null,
                    null,
                    cell(new BigDecimal("2.512351234324832")).withStyle(twoDecimal), //C3
                    cell(false) //D3
            };

            Cell[] s1row4 = new Cell[] {
                    null,
                    null,
                    cell(3.14), //C4
            };

            Cell[] s1row5 = new Cell[] {
                    null,
                    null,
                    cell(new BigDecimal("1.234567890")).withStyle(twoDecimalBuiltin) //C5
            };


            w.withSheet("test", Stream.of(row(s1row1), row(s1row2), row(s1row3), row(s1row4), row(s1row5)));
            //

            //sheet2



            Cell[] s2row1 = new Cell[] {
                    null,
                    cell("World"),
                    cell("Sum")
            };
            Cell[] s2row2 = new Cell[] {
                    cell("Hello").withStyle(borderDiagonal),
                    null,
                    cell(1),
                    cell(dateForSheet2Row1Col3).withStyle(dateFormat)
            };
            Cell[] s2row3 = new Cell[] {
                    null,
                    null,
                    cell(2),
            };
            Cell[] s2row4 = new Cell[] {
                    null,
                    null,
                    formula("SUM(C2:C3)"),
            };
            Cell[] s2row5 = new Cell[] {
                    null,
                    null,
                    formula("C2+C3+C4"),
            };

            w.withSheet("test2", Stream.of(row(s2row1), row(s2row2), row(s2row3), row(s2row4), row(s2row5)));

        }

        if (OUTPUT_FILE) {
            try (FileOutputStream fos = new FileOutputStream("test2.xlsx")) {
                fos.write(baos.toByteArray());
            }
        }

        PoiCrossCheck.checkResult(dateForSheet2Row1Col3, new ByteArrayInputStream(baos.toByteArray()));
    }
}
