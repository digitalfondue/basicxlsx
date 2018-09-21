package ch.digitalfondue.basicxlsx;

import org.w3c.dom.Document;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.stream.Collectors;

class Utils {

    static String readFromResource(String resource) {
        try (InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream(resource);
             BufferedReader b = new BufferedReader(new InputStreamReader(is, StandardCharsets.UTF_8))) {
            return b.lines().collect(Collectors.joining());
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    static Document toDocument(String resource) {
        try (InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream(resource)) {
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            dbFactory.setNamespaceAware(true);
            dbFactory.setExpandEntityReferences(false);
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            return dBuilder.parse(is);
        } catch (ParserConfigurationException | IOException | SAXException e) {
            throw new IllegalStateException(e);
        }
    }

    static String fromDocument(Document doc) {
        try {
            DOMSource domSource = new DOMSource(doc);
            Transformer transformer = TransformerFactory.newInstance().newTransformer();
            StringWriter sw = new StringWriter();
            StreamResult sr = new StreamResult(sw);
            transformer.transform(domSource, sr);
            return sw.toString();
        } catch (TransformerException e) {
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
}
