package ch.digitalfondue.basicxlsx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.math.BigDecimal;

public abstract class Cell {


    abstract Element toElement(Document doc, int row, int column);


    private static Element buildCell(Document doc, String type, int row, int column) {
        Element cell = doc.createElementNS(Utils.NS_SPREADSHEETML_2006_MAIN, "c");
        cell.setAttribute("r", Utils.fromRowColumnToExcelCoordinates(row, column));
        cell.setAttribute("t", type);
        return cell;
    }

    //inline string element
    public static class StringCell extends Cell {

        private final String value;

        public StringCell(String value) {
            this.value = value;
        }

        //http://officeopenxml.com/SScontentOverview.php
        @Override
        Element toElement(Document doc, int row, int column) {

            // <c r="B1" t="inlineStr">
            //  <is>
            //    <t>Name1</t>
            //  </is>
            // </c>
            Element cell = buildCell(doc, "inlineStr", row, column);

            Element is = doc.createElementNS(Utils.NS_SPREADSHEETML_2006_MAIN, "is");
            Element t = doc.createElementNS(Utils.NS_SPREADSHEETML_2006_MAIN, "t");
            t.setTextContent(value);
            is.appendChild(t);
            cell.appendChild(is);

            return cell;
        }
    }

    //number
    public static class NumberCell extends Cell {
        private final BigDecimal number;

        public NumberCell(long number) {
            this.number = BigDecimal.valueOf(number);
        }

        public NumberCell(double number) {
            this.number = BigDecimal.valueOf(number);
        }

        public NumberCell(BigDecimal number) {
            this.number = number;
        }


        //http://officeopenxml.com/SScontentOverview.php
        @Override
        Element toElement(Document doc, int row, int column) {

            // <c r="B2" t="n">
            //  <v>400</v>
            // </c>
            Element cell = buildCell(doc, "n", row, column);
            Element v = doc.createElementNS(Utils.NS_SPREADSHEETML_2006_MAIN, "v");
            v.setTextContent(number.toPlainString());

            cell.appendChild(v);

            return cell;
        }
    }

    //boolean
    public static class BooleanCell extends Cell {
        private final boolean value;

        public BooleanCell(boolean value) {
            this.value = value;
        }

        @Override
        Element toElement(Document doc, int row, int column) {
            // <c r="B2" t="b">
            //  <v>1</v>
            // </c>
            Element cell = buildCell(doc, "b", row, column);
            Element v = doc.createElementNS(Utils.NS_SPREADSHEETML_2006_MAIN, "v");
            v.setTextContent(value ? "1" : "0");
            cell.appendChild(v);
            return cell;
        }
    }
}
