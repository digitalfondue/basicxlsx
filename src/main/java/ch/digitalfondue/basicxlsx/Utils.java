package ch.digitalfondue.basicxlsx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
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
            Transformer transformer = TransformerFactory.newInstance().newTransformer();
            StreamResult sr = new StreamResult(new OutputStreamWriter(os, StandardCharsets.UTF_8));
            transformer.transform(domSource, sr);
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
