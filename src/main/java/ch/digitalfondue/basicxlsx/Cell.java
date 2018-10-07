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
import java.util.function.Function;

/**
 * Represent a cell.
 */
public abstract class Cell {

    abstract Element toElement(Function<String, Element> elementBuilder, int row, int column, int styleId);

    Style style;

    /**
     * Build a cell for xml serialization purpose.
     *
     * @param elementBuilder
     * @param type
     * @param row
     * @param column
     * @param styleId
     * @return
     */
    static Element buildCell(Function<String, Element> elementBuilder, String type, int row, int column, int styleId) {
        Element cell = elementBuilder.apply("c");
        cell.setAttribute("r", Utils.fromRowColumnToExcelCoordinates(row, column));
        cell.setAttribute("t", type);
        cell.setAttribute("s", Integer.toString(styleId));
        return cell;
    }

    /**
     * Set the style to this cell. You can pass null to remove the style if necessary.
     *
     * @param style
     * @return
     */
    public final Cell withStyle(Style style) {
        this.style = style;
        return this;
    }

    /**
     * Get the style associated with the cell, may return null.
     *
     * @return
     */
    public final Style getStyle() {
        return style;
    }


    //
    public static Cell cell(String value) {
        return new StringCell(value);
    }

    public static Cell cell(long value) {
        return new NumberCell(BigDecimal.valueOf(value));
    }

    public static Cell cell(double value) {
        return new NumberCell(BigDecimal.valueOf(value));
    }

    public static Cell cell(BigDecimal value) {
        return new NumberCell(value);
    }

    public static Cell cell(boolean value) {
        return new BooleanCell(value);
    }

    public static Cell cell(Date value) {
        return new DateCell(Utils.getExcelDate(value));
    }

    public static Cell cell(LocalDateTime value) {
        return new DateCell(Utils.getExcelDate(value));
    }

    public static Cell cell(LocalDate value) {
        return new DateCell(Utils.getExcelDate(value));
    }

    public static Cell formula(String formula) {
        return new FormulaCell(formula, null);
    }

    public static Cell formula(String formula, String value) {
        return new FormulaCell(formula, value);
    }
    //

    //inline string element
    /**
     * Cell that contains a string.
     */
    private static class StringCell extends Cell {

        private final String value;

        private StringCell(String value) {
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
    private static class NumberCell extends Cell {
        private final BigDecimal number;

        private NumberCell(BigDecimal number) {
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

    // formula
    private static class FormulaCell extends Cell {
        private final String formula;
        private final String result;

        private FormulaCell(String formula, String result) {
            this.formula = formula;
            this.result = result;
        }

        @Override
        Element toElement(Function<String, Element> elementBuilder, int row, int column, int styleId) {
            //<c r="B9" t="str">
            //<f>SUM(B2:B8)</f>
            //<v>2105</v>
            //</c>
            Element cell = buildCell(elementBuilder, "n", row, column, styleId);
            Element f = elementBuilder.apply("f");
            f.setTextContent(this.formula);
            cell.appendChild(f);

            if (result != null) {
                Element v = elementBuilder.apply("v");
                v.setTextContent(result);
                cell.appendChild(v);
            }
            return cell;
        }
    }

    //boolean

    /**
     * Cell with a boolean value.
     */
    private static class BooleanCell extends Cell {
        private final boolean value;

        private BooleanCell(boolean value) {
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


    private static class DateCell extends Cell {

        private final BigDecimal value;

        DateCell(BigDecimal value) {
            this.value = value;
        }

        @Override
        Element toElement(Function<String, Element> elementBuilder, int row, int column, int styleId) {
            // <c r="B2" t="n">
            //  <v>400</v>
            // </c>
            Element cell = buildCell(elementBuilder, "n", row, column, styleId);
            Element v = elementBuilder.apply("v");
            v.setTextContent(value.toPlainString());
            cell.appendChild(v);
            return cell;
        }
    }
}
