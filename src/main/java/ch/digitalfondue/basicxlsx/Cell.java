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

import org.w3c.dom.Element;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.function.BiFunction;
import java.util.function.Function;

/**
 * Base class for the concrete implementations.
 */
public abstract class Cell {

    abstract Element toElement(Function<String, Element> elementBuilder, int row, int column, int styleId);

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
    /**
     * Cell that contains a string.
     */
    public static class StringCell extends Cell {

        private final String value;

        public StringCell(String value) {
            this.value = value;
        }

        //http://officeopenxml.com/SScontentOverview.php
        @Override
        Element toElement(Function<String, Element> elementBuilder, int row, int column, int styleId) {

            // <c r="B1" t="inlineStr">
            //  <is>
            //    <t>Name1</t>
            //  </is>
            // </c>
            Element cell = buildCell(elementBuilder, "inlineStr", row, column, styleId);

            Element is = elementBuilder.apply("is");
            Element t = elementBuilder.apply("t");
            t.setTextContent(value);
            is.appendChild(t);
            cell.appendChild(is);

            return cell;
        }
    }

    //number

    /**
     * Cell with a numeric value.
     */
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
        Element toElement(Function<String, Element> elementBuilder, int row, int column, int styleId) {

            // <c r="B2" t="n">
            //  <v>400</v>
            // </c>
            Element cell = buildCell(elementBuilder, "n", row, column, styleId);
            Element v = elementBuilder.apply("v");
            v.setTextContent(number.toPlainString());

            cell.appendChild(v);

            return cell;
        }
    }

    //boolean

    /**
     * Cell with a boolean value.
     */
    public static class BooleanCell extends Cell {
        private final boolean value;

        public BooleanCell(boolean value) {
            this.value = value;
        }

        @Override
        Element toElement(Function<String, Element> elementBuilder, int row, int column, int styleId) {
            // <c r="B2" t="b">
            //  <v>1</v>
            // </c>
            Element cell = buildCell(elementBuilder, "b", row, column, styleId);
            Element v = elementBuilder.apply("v");
            v.setTextContent(value ? "1" : "0");
            cell.appendChild(v);
            return cell;
        }
    }


    private static abstract class AbstractDateCell<T> extends Cell {

        private final T value;
        private final Function<T, BigDecimal> converter;

        AbstractDateCell(T value, Function<T, BigDecimal> converter) {
            this.value = value;
            this.converter = converter;
        }

        @Override
        Element toElement(Function<String, Element> elementBuilder, int row, int column, int styleId) {
            // <c r="B2" t="n">
            //  <v>400</v>
            // </c>
            Element cell = buildCell(elementBuilder, "n", row, column, styleId);
            Element v = elementBuilder.apply("v");
            v.setTextContent(converter.apply(value).toPlainString());
            cell.appendChild(v);
            return cell;
        }
    }

    // date

    /**
     * Cell with a date. Please note that a formatting must be provided.
     */
    public static class DateCell extends AbstractDateCell<Date> {
        public DateCell(Date value) {
            super(value, Utils::getExcelDate);
        }
    }

    // date variant
    /**
     * Cell with a date. Please note that a formatting must be provided.
     */
    public static class LocalDateTimeCell extends AbstractDateCell<LocalDateTime> {
        public LocalDateTimeCell(LocalDateTime value) {
            super(value, Utils::getExcelDate);
        }
    }

    /**
     * Cell with a date. Please note that a formatting must be provided.
     */
    public static class LocalDateCell extends AbstractDateCell<LocalDate> {
        public LocalDateCell(LocalDate value) {
            super(value, Utils::getExcelDate);
        }
    }
    //
}
