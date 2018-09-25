package ch.digitalfondue.basicxlsx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.math.BigDecimal;
import java.util.function.BiFunction;
import java.util.function.Function;

public abstract class Cell {

    abstract Element toElement(Function<String, Element> elementBuilder, int row, int column);

    BiFunction<Cell, Style, Style> styleRegistrator;
    Function<Cell, Style> styleRegistry;

    private static Element buildCell(Function<String, Element> elementBuilder, String type, int row, int column, int styleId) {
        Element cell = elementBuilder.apply("c");
        cell.setAttribute("r", Utils.fromRowColumnToExcelCoordinates(row, column));
        cell.setAttribute("t", type);
        cell.setAttribute("s", Integer.toString(styleId));
        return cell;
    }

    public final Cell withStyle(Style style) {
        if (styleRegistrator == null) {
            throw new IllegalStateException("Cell cannot be styled if not registered in a sheet");
        }
        styleRegistrator.apply(this, style);
        return this;
    }

    public final Style getStyle() {
        if (styleRegistry == null) {
            throw new IllegalStateException("Cell cannot be styled if not registered in a sheet");
        }
        return styleRegistry.apply(this);
    }

    //inline string element
    public static class StringCell extends Cell {

        private final String value;

        public StringCell(String value) {
            this.value = value;
        }

        //http://officeopenxml.com/SScontentOverview.php
        @Override
        Element toElement(Function<String, Element> elementBuilder, int row, int column) {

            // <c r="B1" t="inlineStr">
            //  <is>
            //    <t>Name1</t>
            //  </is>
            // </c>
            Element cell = buildCell(elementBuilder, "inlineStr", row, column, 0); //FIXME get the styleId

            Element is = elementBuilder.apply("is");
            Element t = elementBuilder.apply("t");
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
        Element toElement(Function<String, Element> elementBuilder, int row, int column) {

            // <c r="B2" t="n">
            //  <v>400</v>
            // </c>
            Element cell = buildCell(elementBuilder, "n", row, column, 0);
            Element v = elementBuilder.apply("v");
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
        Element toElement(Function<String, Element> elementBuilder, int row, int column) {
            // <c r="B2" t="b">
            //  <v>1</v>
            // </c>
            Element cell = buildCell(elementBuilder, "b", row, column, 0);
            Element v = elementBuilder.apply("v");
            v.setTextContent(value ? "1" : "0");
            cell.appendChild(v);
            return cell;
        }
    }
}
