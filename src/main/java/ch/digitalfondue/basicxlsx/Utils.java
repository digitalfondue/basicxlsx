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
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;
import java.util.TimeZone;
import java.util.function.Function;

class Utils {

    static final String NS_SPREADSHEETML_2006_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    static Function<String, Element> toElementBuilder(Document doc) {
        return (elemName) -> doc.createElementNS(Utils.NS_SPREADSHEETML_2006_MAIN, elemName);
    }

    static Document toDocument(String resource) {
        try (InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream(resource)) {
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
            if(omitXmlPrologue) {
                tf.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
            }
            return  tf;
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
        return sb.append(Integer.toString(row + 1)).toString();
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
}
