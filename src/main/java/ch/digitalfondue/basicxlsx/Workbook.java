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
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class Workbook {

    private final Map<String, Sheet> sheets = new HashMap<>();
    private final List<String> sheetNameOrder = new ArrayList<>();
    private final List<Style> styles = new ArrayList<>();
    final Map<Cell, Style> styledCells = new IdentityHashMap<>();
    final Map<Style, Integer> styleToIdMapping = new IdentityHashMap();

    public Sheet sheet(String name) {
        return sheets.computeIfAbsent(name, sheetName -> {
            sheetNameOrder.add(sheetName);
            return new Sheet(styledCells);
        });
    }

    private int styleIdSupplier(Cell cell) {
        if (styledCells.containsKey(cell)) {
            Integer r = styleToIdMapping.get(styledCells.get(cell));
            return r != null ? r : 0;
        } else {
            return 0;//default id
        }
    }


    public Style.StyleBuilder defineStyle() {
        return Style.define(styles::add);
    }

    public void write(OutputStream os) throws IOException {
        try (ZipOutputStream zos = new ZipOutputStream(os, StandardCharsets.UTF_8)) {

            addFileWithDocument(zos, "[Content_Types].xml", buildContentTypes(sheets.size()));

            addFileWithDocument(zos, "_rels/.rels", buildRels());

            addFileWithDocument(zos, "xl/workbook.xml", buildWorkbook(sheets.size(), sheetNameOrder));

            addFileWithDocument(zos, "xl/styles.xml", buildStyles(styles, styleToIdMapping));

            addFileWithDocument(zos, "xl/_rels/workbook.xml.rels", buildWorkbookRels(sheets.size()));

            for (int i = 0; i < sheets.size(); i++) {
                addFileWithDocument(zos, "xl/worksheets/sheet" + (i + 1) + ".xml", buildSheet(sheets.get(sheetNameOrder.get(i)), this::styleIdSupplier));
            }
        }
    }

    private static void addFileWithDocument(ZipOutputStream zos, String fileName, Document doc) throws IOException {
        zos.putNextEntry(new ZipEntry(fileName));
        Utils.outputDocument(doc, zos);
        zos.closeEntry();
    }

    private static Element getElement(Element element, String name) {
        return (Element) element.getElementsByTagNameNS(Utils.NS_SPREADSHEETML_2006_MAIN, name).item(0);
    }

    private static Element getElement(Document doc, String name) {
        return getElement(doc.getDocumentElement(), name);
    }

    private static void adjustCount(Element element, String childName) {
        element.setAttribute("count", Integer.toString(element.getElementsByTagNameNS(Utils.NS_SPREADSHEETML_2006_MAIN, childName).getLength()));
    }

    private static Document buildStyles(List<Style> styles, Map<Style, Integer> styleToIdMapping) {
        //FIXME implement
        Document doc = Utils.toDocument("ch/digitalfondue/basicxlsx/styles_template.xml");
        Function<String, Element> elementBuilder = Utils.toElementBuilder(doc);

        Element fonts = getElement(doc, "fonts");
        Element cellXfs = getElement(doc, "cellXfs");

        for (Style style : styles) {
            int styleId = style.register(elementBuilder, fonts, cellXfs);
            styleToIdMapping.put(style, styleId);
        }


        //
        adjustCount(fonts, "font");
        adjustCount(cellXfs, "xf");
        //
        return doc;
    }

    private static Document buildRels() {
        return Utils.toDocument("ch/digitalfondue/basicxlsx/rels_template.xml");
    }

    private static Document buildContentTypes(int sheetCount) {

        Document doc = Utils.toDocument("ch/digitalfondue/basicxlsx/content_types_template.xml");
        Element root = doc.getDocumentElement();

        // Override elements for the sheets
        for (int i = 0; i < sheetCount; i++) {
            // <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            Element overrideElem = doc.createElementNS("http://schemas.openxmlformats.org/package/2006/content-types", "Override");
            overrideElem.setAttribute("PartName", "/xl/worksheets/sheet" + (i + 1) + ".xml");
            overrideElem.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            root.appendChild(overrideElem);
        }
        return doc;
    }

    private static Document buildWorkbookRels(int sheetCount) {

        Document doc = Utils.toDocument("ch/digitalfondue/basicxlsx/workbook_rels_template.xml");
        Element root = doc.getDocumentElement();
        // add for each sheet
        // <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="xl/worksheets/sheet1.xml"/>
        for (int i = 0; i < sheetCount; i++) {
            Element rel = doc.createElementNS("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship");
            rel.setAttribute("Id", "rId" + (i + 2)); //rdId1 it's the style.xml file
            rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            rel.setAttribute("Target", "/xl/worksheets/sheet" + (i + 1) + ".xml");
            root.appendChild(rel);
        }
        return doc;
    }

    private static Document buildWorkbook(int sheetCount, List<String> sheetNameOrder) {

        Document doc = Utils.toDocument("ch/digitalfondue/basicxlsx/workbook_template.xml");
        Element root = getElement(doc, "sheets");
        // <sheet name="Table0" sheetId="1" r:id="rId1"/>
        for (int i = 0; i < sheetCount; i++) {
            Element sheet = doc.createElementNS(Utils.NS_SPREADSHEETML_2006_MAIN, "sheet");
            sheet.setAttribute("name", sheetNameOrder.get(i));
            sheet.setAttribute("sheetId", Integer.toString(i + 1));
            sheet.setAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id", "rId" + (i + 2)); //rdId1 it's the style.xml file
            root.appendChild(sheet);
        }
        return doc;
    }

    private static Document buildSheet(Sheet sheet, Function<Cell, Integer> styleIdSupplier) {
        Document doc = Utils.toDocument("ch/digitalfondue/basicxlsx/sheet_template.xml");
        Element cols = getElement(doc, "cols");
        Function<String, Element> elementBuilder = Utils.toElementBuilder(doc);

        final int colsCount = sheet.getMaxCol() + 1;
        for (int i = 0; i < colsCount; i++) {
            Element col = elementBuilder.apply("col");
            col.setAttribute("min", Integer.toString(i + 1));
            col.setAttribute("max", Integer.toString(i + 1));
            cols.appendChild(col);
        }

        Element sheetData = getElement(doc, "sheetData");

        //row
        for (Map.Entry<Integer, SortedMap<Integer, Cell>> rowCells : sheet.cells.entrySet()) {
            Element row = elementBuilder.apply("row");
            row.setAttribute("r", Integer.toString(rowCells.getKey() + 1));

            //column -> cell
            for (Map.Entry<Integer, Cell> colAndCell : rowCells.getValue().entrySet()) {
                Cell cell = colAndCell.getValue();
                int styleId = styleIdSupplier.apply(cell);
                row.appendChild(cell.toElement(elementBuilder, rowCells.getKey(), colAndCell.getKey(), styleId));
            }
            sheetData.appendChild(row);
        }
        return doc;
    }
}
