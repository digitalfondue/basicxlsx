package ch.digitalfondue.basicxlsx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class Workbook {

    private final Map<String, Sheet> sheets = new HashMap<>();
    private final List<String> sheetNameOrder = new ArrayList<>();

    public Sheet openSheet(String name) {
        return sheets.computeIfAbsent(name, sheetName -> {
            sheetNameOrder.add(sheetName);
            return new Sheet(sheetName);
        });
    }

    public void write(OutputStream os) throws IOException {
        try (ZipOutputStream zos = new ZipOutputStream(os, StandardCharsets.UTF_8)) {

            addFileWithContent(zos, "[Content_Types].xml", buildContentTypes(sheets.size()));

            addFileWithContent(zos, "_rels/.rels", buildRels());

            addFileWithContent(zos, "xl/workbook.xml", buildWorkbook(sheets.size(), sheetNameOrder));

            addFileWithContent(zos, "xl/_rels/workbook.xml.rels", buildWorkbookRels(sheets.size()));

            for (int i = 0; i < sheets.size(); i++) {
                addFileWithContent(zos, "xl/worksheets/sheet" + (i + 1) + ".xml", buildSheet(sheets.get(sheetNameOrder.get(i))));
            }
        }
    }

    private static void addFileWithContent(ZipOutputStream zos, String file, String content) throws IOException {
        zos.putNextEntry(new ZipEntry(file));
        zos.write(content.getBytes(StandardCharsets.UTF_8));
        zos.closeEntry();
    }

    private static String buildRels() {
        return Utils.readFromResource("ch/digitalfondue/basicxlsx/rels_template.xml");
    }


    private static String buildContentTypes(int sheetCount) {

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
        return Utils.fromDocument(doc);
    }

    private static String buildWorkbookRels(int sheetCount) {

        Document doc = Utils.toDocument("ch/digitalfondue/basicxlsx/workbook_rels_template.xml");
        Element root = doc.getDocumentElement();
        // add for each sheet
        // <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="xl/worksheets/sheet1.xml"/>
        for (int i = 0; i < sheetCount; i++) {
            Element rel = doc.createElementNS("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship");
            rel.setAttribute("Id", "rId" + (i + 1));
            rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            rel.setAttribute("Target", "/xl/worksheets/sheet" + (i + 1) + ".xml");
            root.appendChild(rel);
        }
        return Utils.fromDocument(doc);
    }

    private static String buildWorkbook(int sheetCount, List<String> sheetNameOrder) {

        Document doc = Utils.toDocument("ch/digitalfondue/basicxlsx/workbook_template.xml");
        Node root = doc.getDocumentElement().getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "sheets").item(0);
        // <sheet name="Table0" sheetId="1" r:id="rId1"/>
        for (int i = 0; i < sheetCount; i++) {
            Element sheet = doc.createElementNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "sheet");
            sheet.setAttribute("name", sheetNameOrder.get(i));
            sheet.setAttribute("sheetId", Integer.toString(i + 1));
            sheet.setAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id", "rId" + (i + 1));
            root.appendChild(sheet);
        }
        return Utils.fromDocument(doc);
    }

    private static String buildSheet(Sheet sheet) {
        System.err.println("Sheet is " + sheet.getName());
        System.err.println("max col is " + sheet.getMaxCol());
        Document doc = Utils.toDocument("ch/digitalfondue/basicxlsx/sheet_template.xml");
        Element cols = (Element) doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "cols").item(0);

        final int colsCount = sheet.getMaxCol() + 1;
        for (int i = 0; i < colsCount; i++) {
            Element col = doc.createElementNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "col");
            col.setAttribute("min", Integer.toString(i + 1));
            col.setAttribute("max", Integer.toString(i + 1));
            cols.appendChild(col);
        }


        System.err.println(Utils.fromDocument(doc));
        return Utils.fromDocument(doc);
    }
}
