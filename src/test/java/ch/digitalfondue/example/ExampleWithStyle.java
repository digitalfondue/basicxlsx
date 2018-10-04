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
package ch.digitalfondue.example;

import ch.digitalfondue.basicxlsx.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExampleWithStyle {

    public static void main(String args[]) throws IOException {
        Workbook w = new Workbook();


        // you must define the styles before
        Style redBGBold = w.defineStyle().bgColor(Style.Color.RED).font().bold(true).build();
        //

        Sheet s = w.sheet("test"); //create a new sheet named test
        s.setValueAt("Hello World", 0, 0).withStyle(redBGBold); //put in "A1" the value "Hello World", set the style to the cell

        try (FileOutputStream fos = new FileOutputStream("test.xlsx")) {
            w.write(fos);
        }
    }
}
