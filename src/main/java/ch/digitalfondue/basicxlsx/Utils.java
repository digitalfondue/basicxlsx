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
package ch.digitalfondue.basicxlsx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.*;
import java.util.function.Function;

class Utils {

    static final String NS_SPREADSHEETML_2006_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    static final Map<String, byte[]> xmlTemplates = Map.of(
            "content_types_template.xml", /* language=XML */ ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "    <Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/>" +
                    "    <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "    <Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                    "    <Override PartName=\"/xl/workbook.xml\"" +
                    "              ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                    "    <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
                    "    <!--" +
                    "    <Override PartName=\"/xl/worksheets/sheet1.xml\"" +
                    "              ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                    "    <Override PartName=\"/xl/worksheets/sheet2.xml\"" +
                    "              ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                    "              -->" +
                    "</Types>").getBytes(StandardCharsets.UTF_8),
            "rels_template.xml", /* language=XML */ ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "    <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\"" +
                    "                  Target=\"xl/workbook.xml\"/>" +
                    "</Relationships>").getBytes(StandardCharsets.UTF_8),
            "sheet_template.xml", /* language=XML */ ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                    "    <sheetViews>" +
                    "        <sheetView workbookViewId=\"0\" tabSelected=\"true\"/>" +
                    "    </sheetViews>" +
                    "    <cols>" +
                    "        <!--<col min=\"1\" max=\"1\"/>-->" +
                    "    </cols>" +
                    "    <sheetData>" +
                    "        <!--<row r=\"1\">" +
                    "            <c r=\"B1\" t=\"inlineStr\">" +
                    "                <is>" +
                    "                    <t>Name1</t>" +
                    "                </is>" +
                    "            </c>" +
                    "        </row>" +
                    "        -->" +
                    "    </sheetData>" +
                    "</worksheet>").getBytes(StandardCharsets.UTF_8),
            "styles_template.xml", /* language=XML */ ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                    "    <!-- http://officeopenxml.com/SSstyles.php -->" +
                    "    <numFmts count=\"1\">" +
                    "        <!-- based from openoffice output -->" +
                    "        <numFmt numFmtId=\"164\" formatCode=\"General\"/>" +
                    "    </numFmts>" +
                    "    <fonts count=\"4\">" +
                    "        <!-- based from https://github.com/mk-j/PHP_XLSXWriter/blob/master/xlsxwriter.class.php#L472 and openoffice output -->" +
                    "        <font>" +
                    "            <sz val=\"10\"/>" +
                    "            <name val=\"Arial\"/>" +
                    "            <family val=\"2\"/>" +
                    "        </font>" +
                    "        <font>" +
                    "            <sz val=\"10\"/>" +
                    "            <name val=\"Arial\"/>" +
                    "            <family val=\"0\"/>" +
                    "        </font>" +
                    "        <font>" +
                    "            <sz val=\"10\"/>" +
                    "            <name val=\"Arial\"/>" +
                    "            <family val=\"0\"/>" +
                    "        </font>" +
                    "        <font>" +
                    "            <sz val=\"10\"/>" +
                    "            <name val=\"Arial\"/>" +
                    "            <family val=\"0\"/>" +
                    "        </font>" +
                    "        <!-- hardcoded test -->" +
                    "        <!-- bold --><!--<font>" +
                    "            <b val=\"true\"/>" +
                    "            <sz val=\"10\"/>" +
                    "            <name val=\"Arial\"/>" +
                    "            <family val=\"2\"/>" +
                    "        </font>-->" +
                    "        <!-- italic --><!--<font>" +
                    "            <i val=\"true\"/>" +
                    "            <sz val=\"10\"/>" +
                    "            <name val=\"Arial\"/>" +
                    "            <family val=\"2\"/>" +
                    "        </font>-->" +
                    "        <!-- -->" +
                    "    </fonts>" +
                    "    <fills count=\"2\">" +
                    "        <!-- based from openoffice output -->" +
                    "        <fill><patternFill patternType=\"none\"/></fill>" +
                    "        <fill><patternFill patternType=\"gray125\"/></fill>" +
                    "    </fills>" +
                    "    <borders count=\"1\">" +
                    "        <!-- based from openoffice output -->" +
                    "        <border diagonalDown=\"false\" diagonalUp=\"false\"><left/><right/><top/><bottom/><diagonal/></border>" +
                    "    </borders>" +
                    "    <!-- -->" +
                    "    <cellStyleXfs count=\"20\"> <!-- based from openoffice output -->" +
                    "        <xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"true\" applyAlignment=\"true\" applyProtection=\"true\">" +
                    "            <alignment horizontal=\"general\" vertical=\"bottom\" textRotation=\"0\" wrapText=\"false\" indent=\"0\" shrinkToFit=\"false\"/>" +
                    "            <protection locked=\"true\" hidden=\"false\"/>" +
                    "        </xf>" +
                    "        <xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"43\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"41\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"44\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"42\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "        <xf numFmtId=\"9\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>" +
                    "    </cellStyleXfs>" +
                    "    <cellXfs count=\"1\">" +
                    "        <!-- based from openoffice output -->" +
                    "        <!-- default -->" +
                    "        <xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"false\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\">" +
                    "            <alignment horizontal=\"general\" vertical=\"bottom\" textRotation=\"0\" wrapText=\"false\" indent=\"0\" shrinkToFit=\"false\"/>" +
                    "            <protection locked=\"true\" hidden=\"false\"/>" +
                    "        </xf>" +
                    "        <!-- hardcoded test -->" +
                    "        <!-- bold --><!--" +
                    "        <xf numFmtId=\"164\" fontId=\"4\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\">" +
                    "            <alignment horizontal=\"general\" vertical=\"bottom\" textRotation=\"0\" wrapText=\"false\" indent=\"0\" shrinkToFit=\"false\"/>" +
                    "            <protection locked=\"true\" hidden=\"false\"/>" +
                    "        </xf>-->" +
                    "        <!-- italic --><!--" +
                    "        <xf numFmtId=\"164\" fontId=\"5\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\">" +
                    "            <alignment horizontal=\"general\" vertical=\"bottom\" textRotation=\"0\" wrapText=\"false\" indent=\"0\" shrinkToFit=\"false\"/>" +
                    "            <protection locked=\"true\" hidden=\"false\"/>" +
                    "        </xf>-->" +
                    "        <!-- -->" +
                    "    </cellXfs>" +
                    "    <cellStyles>" +
                    "        <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" customBuiltin=\"false\"/>" +
                    "        <cellStyle name=\"Comma\" xfId=\"15\" builtinId=\"3\" customBuiltin=\"false\"/>" +
                    "        <cellStyle name=\"Comma [0]\" xfId=\"16\" builtinId=\"6\" customBuiltin=\"false\"/>" +
                    "        <cellStyle name=\"Currency\" xfId=\"17\" builtinId=\"4\" customBuiltin=\"false\"/>" +
                    "        <cellStyle name=\"Currency [0]\" xfId=\"18\" builtinId=\"7\" customBuiltin=\"false\"/>" +
                    "        <cellStyle name=\"Percent\" xfId=\"19\" builtinId=\"5\" customBuiltin=\"false\"/>" +
                    "    </cellStyles>" +
                    "</styleSheet>").getBytes(StandardCharsets.UTF_8),
            "workbook_rels_template.xml", /* language=XML */ ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "    <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
                    "    <!--" +
                    "    <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\"" +
                    "                  Target=\"worksheets/sheet1.xml\"/>" +
                    "    <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\"" +
                    "                  Target=\"worksheets/sheet2.xml\"/>" +
                    "    -->" +
                    "</Relationships>").getBytes(StandardCharsets.UTF_8),
            "workbook_template.xml", /* language=XML */ ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                    "    <!-- https://github.com/jmcnamara/XlsxWriter/blob/b79f2b9ec2027bd1750c15c4612210dd3cff5be2/xlsxwriter/test/workbook/test_workbook03.py -->" +
                    "    <workbookPr defaultThemeVersion=\"124226\"/>" +
                    "    <bookViews>" +
                    "        <workbookView activeTab=\"0\"/>" +
                    "    </bookViews>" +
                    "    <sheets>" +
                    "        <!--" +
                    "        <sheet name=\"Table0\" sheetId=\"1\" r:id=\"rId2\"/>" +
                    "        <sheet name=\"Table1\" sheetId=\"2\" r:id=\"rId3\"/>" +
                    "         -->" +
                    "    </sheets>" +
                    "    <calcPr calcId=\"124519\" fullCalcOnLoad=\"1\"/>" +
                    "</workbook>").getBytes(StandardCharsets.UTF_8)
    );

    static Element elementWithAttr(Function<String, Element> elementBuilder, String name, String attr, String value) {
        Element element = elementBuilder.apply(name);
        element.setAttribute(attr, value);
        return element;
    }

    static Element elementWithVal(Function<String, Element> elementBuilder, String name, String value) {
        return elementWithAttr(elementBuilder, name, "val", value);
    }

    static String formatColor(String color) {
        if (color.startsWith("#")) {
            color = color.substring(1);
        }
        return "FF" + color.toUpperCase(Locale.ENGLISH);
    }

    static Function<String, Element> toElementBuilder(Document doc) {
        return (elemName) -> doc.createElementNS(Utils.NS_SPREADSHEETML_2006_MAIN, elemName);
    }

    static Document toDocument(String resource) {
        try (InputStream is = new ByteArrayInputStream(xmlTemplates.get(resource))) {
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            dbFactory.setNamespaceAware(true);
            dbFactory.setIgnoringComments(true);
            dbFactory.setExpandEntityReferences(false);
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            return dBuilder.parse(is);
        } catch (ParserConfigurationException | IOException | SAXException e) {
            throw new IllegalStateException(e);
        }
    }

    static void outputDocument(Document doc, OutputStream os) {
        try {
            DOMSource domSource = new DOMSource(doc);
            Transformer transformer = getTransformer(false);
            StreamResult sr = new StreamResult(new OutputStreamWriter(os, StandardCharsets.UTF_8));
            transformer.transform(domSource, sr);
        } catch (TransformerException e) {
            throw new IllegalStateException(e);
        }
    }

    static Transformer getTransformer(boolean omitXmlPrologue) {
        try {

            Transformer tf = TransformerFactory.newInstance().newTransformer();
            if (omitXmlPrologue) {
                tf.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
            }
            return tf;
        } catch (TransformerConfigurationException e) {
            throw new IllegalStateException(e);
        }
    }

    //format from the row/column coordinate to the excel one (e.g. B26 or AA24)
    //based from the code of https://github.com/mk-j/PHP_XLSXWriter/blob/master/xlsxwriter.class.php#L720
    static String fromRowColumnToExcelCoordinates(int row, int column) {
        StringBuilder sb = new StringBuilder();
        for (int i = column; i >= 0; i = (i / 26) - 1) {
            sb.insert(0, (char) (i % 26 + 'A'));
        }
        return sb.append(row + 1).toString();
    }


    // code imported and modified from https://github.com/apache/poi/blob/trunk/src/java/org/apache/poi/ss/usermodel/DateUtil.java
    // under the following license
    /* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at
       http://www.apache.org/licenses/LICENSE-2.0
   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
   ==================================================================== */

    private static final BigDecimal BAD_DATE = BigDecimal.valueOf(-1);   // used to specify that date is invalid
    private static final long DAY_MILLISECONDS = (24 * 60 * 60) * 1000L;

    static BigDecimal getExcelDate(Date date) {
        Calendar calStart = getLocaleCalendar();
        calStart.setTime(date);   // If date includes hours, minutes, and seconds, set them to 0
        return internalGetExcelDate(calStart);
    }

    static BigDecimal getExcelDate(LocalDateTime localDateTime) {
        Date d = Date.from(localDateTime.atZone(ZoneId.systemDefault()).toInstant());
        return getExcelDate(d);
    }

    static BigDecimal getExcelDate(LocalDate localDate) {
        Date d = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
        return getExcelDate(d);
    }

    private static Calendar getLocaleCalendar() {
        return Calendar.getInstance(TimeZone.getTimeZone(ZoneId.systemDefault()), Locale.ROOT);
    }

    private static BigDecimal internalGetExcelDate(Calendar date) {
        if ((date.get(Calendar.YEAR) < 1900)) {
            return BAD_DATE;
        }
        // Because of daylight time saving we cannot use
        //     date.getTime() - calStart.getTimeInMillis()
        // as the difference in milliseconds between 00:00 and 04:00
        // can be 3, 4 or 5 hours but Excel expects it to always
        // be 4 hours.
        // E.g. 2004-03-28 04:00 CEST - 2004-03-28 00:00 CET is 3 hours
        // and 2004-10-31 04:00 CET - 2004-10-31 00:00 CEST is 5 hours
        double fraction = (((date.get(Calendar.HOUR_OF_DAY) * 60
                + date.get(Calendar.MINUTE)
        ) * 60 + date.get(Calendar.SECOND)
        ) * 1000 + date.get(Calendar.MILLISECOND)
        ) / (double) DAY_MILLISECONDS;
        Calendar calStart = dayStart(date);

        double value = fraction + absoluteDay(calStart);

        if (value >= 60) {
            value++;
        }

        return new BigDecimal(value);
    }

    private static Calendar dayStart(final Calendar cal) {
        cal.get(Calendar.HOUR_OF_DAY);   // force recalculation of internal fields
        cal.set(Calendar.HOUR_OF_DAY, 0);
        cal.set(Calendar.MINUTE, 0);
        cal.set(Calendar.SECOND, 0);
        cal.set(Calendar.MILLISECOND, 0);
        cal.get(Calendar.HOUR_OF_DAY);   // force recalculation of internal fields
        return cal;
    }

    private static int absoluteDay(Calendar cal) {
        return cal.get(Calendar.DAY_OF_YEAR) + daysInPriorYears(cal.get(Calendar.YEAR));
    }

    private static int daysInPriorYears(int yr) {
        if (yr < 1900) {
            throw new IllegalArgumentException("'year' must be 1900 or greater");
        }

        int yr1 = yr - 1;
        int leapDays = yr1 / 4   // plus julian leap days in prior years
                - yr1 / 100 // minus prior century years
                + yr1 / 400 // plus years divisible by 400
                - 460;      // leap days in previous 1900 years

        return 365 * (yr - 1900) + leapDays;
    }

    // see https://support.microsoft.com/en-us/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9
    // rules are:
    // no: - / \ ? * : [ ]
    // can't begin or end with '
    // can't be named "History"
    // can't contain more than 31 characters.
    static String convertToExcelCompatibleWorksheetName(String name) {
        var res = name.trim();
        if ("history".equalsIgnoreCase(res)) {
            res = res + "1";
        }

        for (var c : new char[]{'-', '/', '\\', '?', '*', ':', '[', ']'}) {
            res = res.replace(c, '_');
        }
        if (res.startsWith("'")) {
            res = res.substring(1);
        }
        if (res.endsWith("'")) {
            res = res.substring(0, res.length() - 1);
        }

        if (res.isEmpty()) {
            res = "1";
        }
        if (res.length() > 31) {
            res = res.substring(0, 31);
        }
        return res;
    }
}
