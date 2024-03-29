/*
 * Copyright © 2018-2024 digitalfondue (info@digitalfondue.ch)
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
package ch.digitalfondue.example;

import ch.digitalfondue.basicxlsx.Cell;
import ch.digitalfondue.basicxlsx.StreamingWorkbook;
import ch.digitalfondue.basicxlsx.StreamingWorkbook.Row;
import ch.digitalfondue.basicxlsx.Style;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

public class ExampleStreamingWithStyle {

    public static void main(String[] args) throws IOException {

        try (FileOutputStream fos = new FileOutputStream("test.xlsx");
             StreamingWorkbook w = new StreamingWorkbook(fos)) { //<- create a StreamingWorkbook: it require the outputstream

            // you must define the styles before
            Style redBGBold = w.defineStyle().bgColor(Style.Color.RED).font().bold(true).build();
            //

            Cell[] row1 = new Cell[] {Cell.cell("Hello World").withStyle(redBGBold)};
            Stream<Row> rows = Stream.of(StreamingWorkbook.row(row1));


            w.withSheet("test", rows); //write a new sheet named "test" with the stream of rows
        }
    }
}
