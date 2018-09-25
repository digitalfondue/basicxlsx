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
import java.util.Locale;
import java.util.Objects;
import java.util.function.Function;

//based on https://xlsxwriter.readthedocs.io/format.html
public class Style {

    private final FontBuilder fontBuilder;

    private Style(FontBuilder fontBuilder) {
        this.fontBuilder = fontBuilder;
    }


    private static Element elementWithAttr(Function<String, Element> elementBuilder, String name, String attr, String value) {
        Element element = elementBuilder.apply(name);
        element.setAttribute(attr, value);
        return element;
    }

    private static Element elementWithVal(Function<String, Element> elementBuilder, String name, String value) {
        return elementWithAttr(elementBuilder, name, "val", value);
    }

    int register(Function<String, Element> elementBuilder, Element fonts, Element cellXfs) {

        int fontId = 0;

        if (fontBuilder != null) {
            // <font>
            //   <b val="true"/> //<- bold
            //   <i val="true"/> //<- italic
            //   <sz val="10"/> // <- size
            //   <name val="Arial"/> <- font name
            //   <family val="2"/>
            // </font>
            //

            Element font = elementBuilder.apply("font");

            if (fontBuilder.bold) {
                font.appendChild(elementWithVal(elementBuilder, "b", "true"));
            }

            if (fontBuilder.italic) {
                font.appendChild(elementWithVal(elementBuilder, "i", "true"));
            }

            if (fontBuilder.color != null) {
                String color = fontBuilder.color;
                if (color.startsWith("#")) {
                    color = color.substring(1);
                }
                font.appendChild(elementWithAttr(elementBuilder, "color", "rgb", "FF" + color.toUpperCase(Locale.ENGLISH)));
            }

            font.appendChild(elementWithVal(elementBuilder, "sz", fontBuilder.size.toPlainString()));
            font.appendChild(elementWithVal(elementBuilder, "name", fontBuilder.name));
            font.appendChild(elementWithVal(elementBuilder, "family", "2")); //<- hardcoded, what it is?


            fonts.appendChild(font);

            fontId = fonts.getElementsByTagNameNS(Utils.NS_SPREADSHEETML_2006_MAIN, "font").getLength() - 1;
        }

        // add the "xf" element
        // <xf numFmtId="164" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false">
        //  <alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/>
        //  <protection locked="true" hidden="false"/>
        // </xf>
        Element xf = elementBuilder.apply("xf");
        xf.setAttribute("numFmtId", "164");
        xf.setAttribute("fontId", Integer.toString(fontId));
        xf.setAttribute("fillId", "0");
        xf.setAttribute("borderId", "0");
        xf.setAttribute("xfId", "0");
        xf.setAttribute("applyFont", "true");
        xf.setAttribute("applyBorder", "false");
        xf.setAttribute("applyAlignment", "false");
        xf.setAttribute("applyProtection", "false");
        Element alignment = elementBuilder.apply("alignment");
        alignment.setAttribute("horizontal", "general");
        alignment.setAttribute("vertical", "bottom");
        alignment.setAttribute("textRotation", "0");
        alignment.setAttribute("wrapText", "false");
        alignment.setAttribute("indent", "0");
        alignment.setAttribute("shrinkToFit", "false");
        xf.appendChild(alignment);
        Element protection = elementBuilder.apply("protection");
        protection.setAttribute("locked", "true");
        protection.setAttribute("hidden", "false");
        xf.appendChild(protection);
        //
        cellXfs.appendChild(xf);

        return cellXfs.getElementsByTagNameNS(Utils.NS_SPREADSHEETML_2006_MAIN, "xf").getLength() - 1;
    }

    public static class StyleBuilder {

        private final Function<Style, Boolean> register;
        private FontBuilder fontBuilder;

        private StyleBuilder(Function<Style, Boolean> register) {
            this.register = register;
        }

        public FontBuilder font() {
            if (fontBuilder == null) {
                fontBuilder = new FontBuilder(this);
            }
            return fontBuilder;
        }

        public Style build() {
            Style s = new Style(fontBuilder);
            register.apply(s);
            return s;
        }
    }

    public static class FontBuilder {
        String name = "Arial";
        BigDecimal size = BigDecimal.TEN;//default
        String color;
        boolean bold;
        boolean italic;
        FontUnderlineStyle fontUnderlineStyle;
        boolean strikeOut;

        private final StyleBuilder styleBuilder;

        private FontBuilder(StyleBuilder styleBuilder) {
            this.styleBuilder = styleBuilder;
        }

        public StyleBuilder and() {
            return styleBuilder;
        }

        public Style build() {
            return styleBuilder.build();
        }

        public FontBuilder bold(boolean bold) {
            this.bold = bold;
            return this;
        }

        public FontBuilder italic(boolean italic) {
            this.italic = italic;
            return this;
        }

        public FontBuilder name(String name) {
            Objects.requireNonNull(name);
            this.name = name;
            return this;
        }

        /**
         * Set the font with a given color.
         *
         * @param color: format is "#ffcc00" or simply "ffcc00"
         * @return
         */
        public FontBuilder color(String color) {
            this.color = color;
            return this;
        }

        public FontBuilder color(Color color) {
            return color(color.color);
        }

        public FontBuilder size(int size) {
            this.size = BigDecimal.valueOf(size);
            return this;
        }

        public FontBuilder size(double size) {
            this.size = BigDecimal.valueOf(size);
            return this;
        }

        public FontBuilder size(BigDecimal size) {
            this.size = size;
            return this;
        }
    }

    public enum FontUnderlineStyle {
        NONE, SINGLE, DOUBLE, SINGLE_ACCOUNTING_UNDERLINE, DOUBLE_ACCOUNTING_UNDERLINE;
    }

    //color list imported from https://github.com/jmcnamara/XlsxWriter/blob/master/xlsxwriter/format.py#L959
    public enum Color {
        BLACK("#000000"),
        BLUE("#0000FF"),
        BROWN("#800000"),
        CYAN("#00FFFF"),
        GRAY("#808080"),
        GREEN("#008000"),
        LIME("#00FF00"),
        MAGENTA("#FF00FF"),
        NAVY("#000080"),
        ORANGE("#FF6600"),
        PINK("#FF00FF"),
        PURPLE("#800080"),
        RED("#FF0000"),
        SILVER("#C0C0C0"),
        WHITE("#FFFFFF"),
        YELLOW("#FFFF00");

        private final String color;

        Color(String color) {
            this.color = color;
        }
    }


    static StyleBuilder define(Function<Style, Boolean> register) {
        return new StyleBuilder(register);
    }
}
