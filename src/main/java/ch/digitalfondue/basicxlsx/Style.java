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

/**
 * Represent a style that can be applied to cells.
 */
public class Style {

    private final String numericFormat;
    private final Integer numericFormatIndex;
    private final FontDesc fontDesc;
    private final String bgColor;
    private final String fgColor;
    private final Integer rotation;
    private final Pattern pattern;
    private final ReadingOrder readingOrder;

    Style(String numericFormat, Integer numericFormatIndex, String bgColor, String fgColor, Pattern pattern, Integer rotation, ReadingOrder readingOrder, FontDesc fontDesc) {
        this.numericFormat = numericFormat;
        this.numericFormatIndex = numericFormatIndex;
        this.bgColor = bgColor;
        this.fgColor = fgColor;
        this.pattern = pattern;
        this.rotation = rotation;
        this.readingOrder = readingOrder;
        this.fontDesc = fontDesc;
    }


    private static Element elementWithAttr(Function<String, Element> elementBuilder, String name, String attr, String value) {
        Element element = elementBuilder.apply(name);
        element.setAttribute(attr, value);
        return element;
    }

    private static Element elementWithVal(Function<String, Element> elementBuilder, String name, String value) {
        return elementWithAttr(elementBuilder, name, "val", value);
    }

    private static String formatColor(String color) {
        if (color.startsWith("#")) {
            color = color.substring(1);
        }
        return "FF" + color.toUpperCase(Locale.ENGLISH);
    }

    int register(Function<String, Element> elementBuilder, Element fonts, Element cellXfs, Element numFmts, Element fills) {

        int fontId = 0;
        int numFmtId = 164;//default value
        int fillId = 0;

        if (numericFormatIndex != null) {
            numFmtId = numericFormatIndex; //builtin formatting
        } else if (numericFormat != null) {
            Element numFmt = elementWithAttr(elementBuilder, "numFmt", "formatCode", numericFormat);
            int count = numFmts.getElementsByTagNameNS(Utils.NS_SPREADSHEETML_2006_MAIN, "numFmt").getLength();
            numFmtId = 164 + count; //custom formatting
            numFmt.setAttribute("numFmtId", Integer.toString(numFmtId));
            numFmts.appendChild(numFmt);
        }

        if (bgColor != null || fgColor != null || pattern != null) {

            //<fill>
            //  <patternFill patternType="solid">
            //    <fgColor rgb="FFFFEB9C"/>
            //    <bgColor rgb="FFFFEB9C"/>
            //  </patternFill/>
            //</fill>

            Element fill = elementBuilder.apply("fill");
            fillId = fills.getElementsByTagNameNS(Utils.NS_SPREADSHEETML_2006_MAIN, "fill").getLength();

            Element patternFill = elementWithAttr(elementBuilder, "patternFill", "patternType", pattern == null ? "solid" : pattern.toXmlValue());
            fill.appendChild(patternFill);

            // if bgColor is defined but not fgColor, we must create fgColor too
            if (fgColor != null || bgColor != null) {
                patternFill.appendChild(elementWithAttr(elementBuilder, "fgColor", "rgb", formatColor(fgColor != null ? fgColor : bgColor)));
                if (bgColor != null) {
                    patternFill.appendChild(elementWithAttr(elementBuilder, "bgColor", "rgb", formatColor(bgColor)));
                }

            }

            fills.appendChild(fill);
        }

        if (fontDesc != null) {
            // <font>
            //   <b val="true"/> //<- bold
            //   <i val="true"/> //<- italic
            //   <sz val="10"/> // <- size
            //   <name val="Arial"/> <- font name
            //   <family val="2"/>
            // </font>
            //

            Element font = elementBuilder.apply("font");

            if (fontDesc.bold) {
                font.appendChild(elementWithVal(elementBuilder, "b", "true"));
            }

            if (fontDesc.italic) {
                font.appendChild(elementWithVal(elementBuilder, "i", "true"));
            }

            if (fontDesc.color != null) {
                font.appendChild(elementWithAttr(elementBuilder, "color", "rgb", formatColor(fontDesc.color)));
            }

            if (fontDesc.strikeOut) {
                font.appendChild(elementBuilder.apply("strike"));
            }

            FontUnderlineStyle underline = fontDesc.fontUnderlineStyle;
            if (underline != null && underline.hasUElement) {
                if (underline.hasValAttribute) {
                    font.appendChild(elementWithVal(elementBuilder, "u", underline.val));
                } else {
                    font.appendChild(elementBuilder.apply("u"));
                }
            }

            font.appendChild(elementWithVal(elementBuilder, "sz", fontDesc.size.toPlainString()));
            font.appendChild(elementWithVal(elementBuilder, "name", fontDesc.name));
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
        xf.setAttribute("numFmtId", Integer.toString(numFmtId));
        xf.setAttribute("fontId", Integer.toString(fontId));
        xf.setAttribute("fillId", Integer.toString(fillId));
        xf.setAttribute("borderId", "0");
        xf.setAttribute("xfId", "0");
        xf.setAttribute("applyFont", "true");
        xf.setAttribute("applyBorder", "false");
        xf.setAttribute("applyAlignment", readingOrder != null ? "true" : "false");
        xf.setAttribute("applyProtection", "false");
        Element alignment = elementBuilder.apply("alignment");
        alignment.setAttribute("horizontal", "general");
        alignment.setAttribute("vertical", "bottom");
        alignment.setAttribute("textRotation", rotation == null ? "0" : Integer.toString(rotation));
        alignment.setAttribute("wrapText", "false");
        alignment.setAttribute("indent", "0");
        alignment.setAttribute("shrinkToFit", "false");
        if (readingOrder != null) {
            alignment.setAttribute("readingOrder", Integer.toString(readingOrder.val));
        }
        xf.appendChild(alignment);
        Element protection = elementBuilder.apply("protection");
        protection.setAttribute("locked", "true");
        protection.setAttribute("hidden", "false");
        xf.appendChild(protection);
        //
        cellXfs.appendChild(xf);

        return cellXfs.getElementsByTagNameNS(Utils.NS_SPREADSHEETML_2006_MAIN, "xf").getLength() - 1;
    }

    /**
     * Style builder. Use it to define a new Style.
     */
    public static class StyleBuilder {

        private final Function<Style, Boolean> register;
        private FontBuilder fontBuilder;
        private String numericFormat;
        private Integer numericFormatIndex;

        private String bgColor;
        private String fgColor;
        private Pattern pattern;

        private Integer rotation;
        private ReadingOrder readingOrder;

        private StyleBuilder(Function<Style, Boolean> register) {
            this.register = register;
        }

        /**
         * Font related options
         *
         * @return a FontBuilder
         */
        public FontBuilder font() {
            if (fontBuilder == null) {
                fontBuilder = new FontBuilder(this);
            }
            return fontBuilder;
        }

        /**
         * <p>Use the built in numeric format.</p>
         *
         *       0, "General"<br>
         *       1, "0"<br>
         *       2, "0.00"<br>
         *       3, "#,##0"<br>
         *       4, "#,##0.00"<br>
         *       5, "$#,##0_);($#,##0)"<br>
         *       6, "$#,##0_);[Red]($#,##0)"<br>
         *       7, "$#,##0.00);($#,##0.00)"<br>
         *       8, "$#,##0.00_);[Red]($#,##0.00)"<br>
         *       9, "0%"<br>
         *       0xa, "0.00%"<br>
         *       0xb, "0.00E+00"<br>
         *       0xc, "# ?/?"<br>
         *       0xd, "# ??/??"<br>
         *       0xe, "m/d/yy"<br>
         *       0xf, "d-mmm-yy"<br>
         *       0x10, "d-mmm"<br>
         *       0x11, "mmm-yy"<br>
         *       0x12, "h:mm AM/PM"<br>
         *       0x13, "h:mm:ss AM/PM"<br>
         *       0x14, "h:mm"<br>
         *       0x15, "h:mm:ss"<br>
         *       0x16, "m/d/yy h:mm"<br>
         *<p>
         *       // 0x17 - 0x24 reserved for international and undocumented
         *       0x25, "#,##0_);(#,##0)"<br>
         *       0x26, "#,##0_);[Red](#,##0)"<br>
         *       0x27, "#,##0.00_);(#,##0.00)"<br>
         *       0x28, "#,##0.00_);[Red](#,##0.00)"<br>
         *       0x29, "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"<br>
         *       0x2a, "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)"<br>
         *       0x2b, "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"<br>
         *       0x2c, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"<br>
         *       0x2d, "mm:ss"<br>
         *       0x2e, "[h]:mm:ss"<br>
         *       0x2f, "mm:ss.0"<br>
         *       0x30, "##0.0E+0"<br>
         *       0x31, "@" - This is text format.<br>
         *       0x31  "text" - Alias for "@"<br>
         *
         * @param idx
         * @return
         */
        public StyleBuilder numericFormat(int idx) {
            if (idx < 0 || (idx >= 0x17 && idx <= 0x24) || idx > 0x31) {
                throw new IllegalArgumentException("idx must be between [0,0x16] - [0x25,0x31]");
            }
            this.numericFormatIndex = idx;
            return this;
        }

        /**
         * Set a numeric format.
         *
         * @param numericFormat
         * @return
         */
        public StyleBuilder numericFormat(String numericFormat) {
            this.numericFormat = numericFormat;
            return this;
        }

        /**
         * Set the background color.
         *
         * @param bgColor
         * @return
         */
        public StyleBuilder bgColor(String bgColor) {
            this.bgColor = bgColor;
            return this;
        }

        /**
         * Set the background color.
         *
         * @param bgColor
         * @return
         */
        public StyleBuilder bgColor(Color bgColor) {
            this.bgColor = bgColor.color;
            return this;
        }

        /**
         * Set the foreground color.
         *
         * @param fgColor
         * @return
         */
        public StyleBuilder fgColor(String fgColor) {
            this.fgColor = fgColor;
            return this;
        }

        /**
         * Set the foreground color.
         *
         * @param fgColor
         * @return
         */
        public StyleBuilder fgColor(Color fgColor) {
            this.fgColor = fgColor.color;
            return this;
        }

        /**
         * Define the Pattern of the cell.
         *
         * @param pattern
         * @return
         */
        public StyleBuilder pattern(Pattern pattern) {
            this.pattern = pattern;
            return this;
        }

        /**
         * Rotation, valid values: from -90 to 90 and 270
         *
         * @param rotation
         * @return
         */
        public StyleBuilder rotation(int rotation) {
            if (rotation < -90 || (rotation > 90 && rotation != 270)) {
                throw new IllegalArgumentException("accepted value for rotation: from -90 to 90 and 270, passed value: " + rotation);
            }
            this.rotation = rotation;
            return this;
        }

        /**
         * Define the reading order: it may be left to right (LTR) or right to left (RTL).
         *
         * @param readingOrder
         * @return
         */
        public StyleBuilder readingOrder(ReadingOrder readingOrder) {
            this.readingOrder = readingOrder;
            return this;
        }

        /**
         * Generate the style.
         *
         * @return
         */
        public Style build() {
            FontDesc fd = fontBuilder != null ? new FontDesc(fontBuilder.name, fontBuilder.size, fontBuilder.color, fontBuilder.bold,
                    fontBuilder.italic, fontBuilder.fontUnderlineStyle, fontBuilder.strikeOut) : null;
            Style s = new Style(numericFormat, numericFormatIndex, bgColor, fgColor, pattern, rotation, readingOrder, fd);
            register.apply(s);
            return s;
        }
    }

    private static class FontDesc {
        final String name;
        final BigDecimal size;
        final String color;
        final boolean bold;
        final boolean italic;
        final FontUnderlineStyle fontUnderlineStyle;
        final boolean strikeOut;

        FontDesc(String name, BigDecimal size, String color, boolean bold, boolean italic, FontUnderlineStyle fontUnderlineStyle, boolean strikeOut) {
            this.name = name;
            this.size = size;
            this.color = color;
            this.bold = bold;
            this.italic = italic;
            this.fontUnderlineStyle = fontUnderlineStyle;
            this.strikeOut = strikeOut;
        }

    }

    /**
     * Font specific options.
     */
    public static class FontBuilder {
        String name = "Calibri";
        BigDecimal size = BigDecimal.valueOf(11);//default
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

        public FontBuilder underline(FontUnderlineStyle fontUnderlineStyle) {
            Objects.requireNonNull(fontUnderlineStyle);
            this.fontUnderlineStyle = fontUnderlineStyle;
            return this;
        }

        public FontBuilder strikeOut(boolean strikeOut) {
            this.strikeOut = strikeOut;
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
        NONE(null, false, false),
        SINGLE(null, false, true),
        DOUBLE("double"),
        SINGLE_ACCOUNTING_UNDERLINE("singleAccounting"),
        DOUBLE_ACCOUNTING_UNDERLINE("doubleAccounting");

        final String val;
        final boolean hasValAttribute;
        final boolean hasUElement;

        FontUnderlineStyle(String val) {
            this(val, true, true);
        }

        FontUnderlineStyle(String val, boolean hasValAttribute, boolean hasUElement) {
            this.val = val;
            this.hasValAttribute = hasValAttribute;
            this.hasUElement = hasUElement;
        }
    }

    /**
     * Reading order.
     */
    public enum ReadingOrder {
        /**
         * Left to right.
         */
        LTR(1),

        /**
         * Right to left.
         */
        RTL(2);

        final int val;

        ReadingOrder(int val) {
            this.val = val;
        }
    }

    //color list imported from https://github.com/jmcnamara/XlsxWriter/blob/master/xlsxwriter/format.py#L959

    /**
     * Some predefined colors
     */
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


    //Pattern https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.patternvalues?view=openxml-2.8.1
    public enum Pattern {
        DARK_DOWN,
        DARK_GRAY,
        DARK_GRID,
        DARK_HORIZONTAL,
        DARK_TRELLIS,
        DARK_UP,
        DARK_VERTICAL,
        GRAY0625,
        GRAY125,
        LIGHT_DOWN,
        LIGHT_GRAY,
        LIGHT_GRID,
        LIGHT_HORIZONTAL,
        LIGHT_TRELLIS,
        LIGHT_UP,
        LIGHT_VERTICAL,
        MEDIUM_GRAY,
        NONE,
        SOLID;

        String toXmlValue() {
            String s = name().toLowerCase(Locale.ROOT);
            int idxSeparator = s.indexOf('_');
            if (idxSeparator == -1) {
                return s;
            }
            StringBuilder sb = new StringBuilder(s);
            sb.replace(idxSeparator, idxSeparator + 2,
                    String.valueOf(Character.toUpperCase(s.charAt(idxSeparator + 1))));
            return sb.toString();
        }
    }


    static StyleBuilder define(Function<Style, Boolean> register) {
        return new StyleBuilder(register);
    }
}
