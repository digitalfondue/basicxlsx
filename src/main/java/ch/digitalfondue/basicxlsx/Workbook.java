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

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.function.Function;
import java.util.zip.ZipOutputStream;

/**
 * <p>Represent a xlsx workbook. It's the main entry point for generating a xlsx file.</p>
 * <p>This workbook keep all the cells data in memory. See {@link StreamingWorkbook} as an alternative.</p>
 */
public class Workbook extends AbstractWorkbook {

    private final Map<String, Sheet> sheets = new LinkedHashMap<>();

    /**
     * Open or create a new sheet.
     *
     * @param name
     * @return a sheet
     */
    public Sheet sheet(String name) {
        return sheets.computeIfAbsent(name, sheetName -> new Sheet());
    }

    /**
     * Write the current worksheet to the output stream.
     *
     * @param os
     * @throws IOException
     */
    public void write(OutputStream os) throws IOException {
        try (ZipOutputStream zos = new ZipOutputStream(os, StandardCharsets.UTF_8)) {

            List<String> sheetNameOrder = new ArrayList<>(sheets.keySet());

            writeMetadataDocuments(zos, sheetNameOrder);
            commitAndWriteStyleMetadata(zos, styles, styleToIdMapping);

            for (int i = 0; i < sheets.size(); i++) {
                addFileWithDocument(zos, "xl/worksheets/sheet" + (i + 1) + ".xml", buildSheet(sheets.get(sheetNameOrder.get(i)), this::styleIdSupplier));
            }
        }
    }

    private static Document buildSheet(Sheet sheet, Function<Cell, Integer> styleIdSupplier) {
        Document doc = Utils.toDocument("ch/digitalfondue/basicxlsx/sheet_template.xml");

        Function<String, Element> elementBuilder = Utils.toElementBuilder(doc);

        //
        if (sheet.readingOrder != null) {
          Element sheetView = getElement(doc, "sheetView");
          sheetView.setAttribute("rightToLeft", Boolean.toString(sheet.readingOrder == Style.ReadingOrder.RTL));
        }
        //

        //TODO check this. It seems not mandatory, seems to be used for sizing the column?
        Element cols = getElement(doc, "cols");
        final int colsCount = sheet.getMaxCol() + 1;
        for (int i = 0; i < colsCount; i++) {
            Element col = elementBuilder.apply("col");
            col.setAttribute("min", Integer.toString(i + 1));
            col.setAttribute("max", Integer.toString(i + 1));

            if (sheet.columnWidth.containsKey(i)) {
                col.setAttribute("customWidth", "true");
                col.setAttribute("width", Double.toString(sheet.columnWidth.get(i)));
            }

            cols.appendChild(col);
        }
        //

        Element sheetData = getElement(doc, "sheetData");

        //row
        for (Map.Entry<Integer, SortedMap<Integer, Cell>> rowCells : sheet.cells.entrySet()) {
            Element row = elementBuilder.apply("row");

            int rowIndex = rowCells.getKey();

            row.setAttribute("r", Integer.toString(rowIndex + 1));

            if (sheet.rowHeight.containsKey(rowIndex)) {
                row.setAttribute("customHeight", "true");
                row.setAttribute("ht", Double.toString(sheet.rowHeight.get(rowIndex)));
            }

            //column -> cell
            for (Map.Entry<Integer, Cell> colAndCell : rowCells.getValue().entrySet()) {
                Cell cell = colAndCell.getValue();
                int styleId = styleIdSupplier.apply(cell);
                row.appendChild(cell.toElement(elementBuilder, rowIndex, colAndCell.getKey(), styleId));
            }
            sheetData.appendChild(row);
        }
        return doc;
    }
}
