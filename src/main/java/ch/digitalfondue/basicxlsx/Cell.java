package ch.digitalfondue.basicxlsx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

public abstract class Cell {

    final String value;

    public Cell(String value) {
        this.value = value;
    }

    abstract Element toElement(Document doc, int row, int column);


    //inline string element
    public static class StringCell extends Cell {
        public StringCell(String value) {
            super(value);
        }

        @Override
        Element toElement(Document doc, int row, int column) {

            // <c r="B1" t="inlineStr">
            //  <is>
            //    <t>Name1</t>
            //  </is>
            // </c>
            Element cell = doc.createElementNS(Utils.NS_SPREADSHEETML_2006_MAIN, "c");
            cell.setAttribute("r", Utils.fromRowColumnToExcelCoordinates(row, column));
            cell.setAttribute("t", "inlineStr");

            Element is = doc.createElementNS(Utils.NS_SPREADSHEETML_2006_MAIN, "is");
            Element t = doc.createElementNS(Utils.NS_SPREADSHEETML_2006_MAIN, "t");
            t.setTextContent(value);
            is.appendChild(t);
            cell.appendChild(is);

            return cell;
        }
    }
}
