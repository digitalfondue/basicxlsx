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
import java.util.ArrayList;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

class AbstractWorkbook {

    final List<Style> styles = new ArrayList<>();
    final Map<Style, Integer> styleToIdMapping = new IdentityHashMap();

    int styleIdSupplier(Cell cell) {
        if (cell.style != null) {
            Integer r = styleToIdMapping.get(cell.style);
            return r != null ? r : 0;
        } else {
            return 0;//default id
        }
    }


    /**
     * Define a new style.
     *
     * @return a style builder
     */
    public Style.StyleBuilder defineStyle() {
        return Style.define(styles::add);
    }

    static void addFileWithDocument(ZipOutputStream zos, String fileName, Document doc) throws IOException {
        zos.putNextEntry(new ZipEntry(fileName));
        Utils.outputDocument(doc, zos);
        zos.closeEntry();
    }


    static Element getElement(Element element, String name) {
        return (Element) element.getElementsByTagNameNS(Utils.NS_SPREADSHEETML_2006_MAIN, name).item(0);
    }

    static Element getElement(Document doc, String name) {
        return getElement(doc.getDocumentElement(), name);
    }


    private static void adjustCount(Element element, String childName) {
        element.setAttribute("count", Integer.toString(element.getElementsByTagNameNS(Utils.NS_SPREADSHEETML_2006_MAIN, childName).getLength()));
    }

    void commitAndWriteStyleMetadata(ZipOutputStream zos, List<Style> styles, Map<Style, Integer> styleToIdMapping) throws IOException {
        Document doc = Utils.toDocument("ch/digitalfondue/basicxlsx/styles_template.xml");
        Function<String, Element> elementBuilder = Utils.toElementBuilder(doc);

        Element fonts = getElement(doc, "fonts");
        Element cellXfs = getElement(doc, "cellXfs");
        Element numFmts = getElement(doc, "numFmts");
        Element fills = getElement(doc, "fills");
        Element borders = getElement(doc, "borders");

        for (Style style : styles) {
            int styleId = style.register(elementBuilder, fonts, cellXfs, numFmts, fills, borders);
            styleToIdMapping.put(style, styleId);
        }
        //
        adjustCount(fonts, "font");
        adjustCount(cellXfs, "xf");
        adjustCount(numFmts, "numFmt");
        adjustCount(fills, "fill");
        adjustCount(borders, "border");

        addFileWithDocument(zos, "xl/styles.xml", doc);
    }

    static void writeMetadataDocuments(ZipOutputStream zos,
                                       List<String> sheetNameOrder) throws IOException {
        int sheetCount = sheetNameOrder.size();
        addFileWithDocument(zos, "[Content_Types].xml", buildContentTypes(sheetCount));
        addFileWithDocument(zos, "_rels/.rels", buildRels());
        addFileWithDocument(zos, "xl/workbook.xml", buildWorkbook(sheetCount, sheetNameOrder));
        addFileWithDocument(zos, "xl/_rels/workbook.xml.rels", buildWorkbookRels(sheetCount));
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

    private static Document buildRels() {
        return Utils.toDocument("ch/digitalfondue/basicxlsx/rels_template.xml");
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
}
