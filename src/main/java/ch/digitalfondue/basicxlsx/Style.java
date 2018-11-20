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

import static ch.digitalfondue.basicxlsx.Utils.*;

import org.w3c.dom.Element;

import java.math.BigDecimal;
import java.util.EnumMap;
import java.util.Locale;
import java.util.Map;
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
    private final VerticalAlignment verticalAlignment;
    private final HorizontalAlignment horizontalAlignment;
    //
    private final LineStyle diagonalLineStyle;
    private final String diagonalColor;
    private final DiagonalStyle diagonalStyle;
    //
    private final BorderDesc borderDesc;
    //

    Style(String numericFormat, Integer numericFormatIndex, String bgColor, String fgColor, Pattern pattern, Integer rotation,
          LineStyle diagonalLineStyle, String diagonalColor, DiagonalStyle diagonalStyle,
          ReadingOrder readingOrder,
          VerticalAlignment verticalAlignment,
          HorizontalAlignment horizontalAlignment,
          FontDesc fontDesc,
          BorderDesc borderDesc) {
        this.numericFormat = numericFormat;
        this.numericFormatIndex = numericFormatIndex;
        this.bgColor = bgColor;
        this.fgColor = fgColor;
        this.pattern = pattern;
        this.rotation = rotation;
        //
        this.diagonalLineStyle = diagonalLineStyle;
        this.diagonalColor = diagonalColor;
        this.diagonalStyle = diagonalStyle;
        //
        this.readingOrder = readingOrder;
        this.verticalAlignment = verticalAlignment;
        this.horizontalAlignment = horizontalAlignment;
        this.fontDesc = fontDesc;
        //
        this.borderDesc = borderDesc;
    }

    FontDesc getFontDesc() {
        return fontDesc;
    }

    Integer getRotation() {
        return rotation;
    }

    int register(Function<String, Element> elementBuilder, Element fonts, Element cellXfs, Element numFmts, Element fills, Element borders) {

        int fontId = 0;
        int numFmtId = 164;//default value
        int fillId = 0;
        int borderId = 0;

        if (numericFormatIndex != null) {
            numFmtId = numericFormatIndex; //builtin formatting
        } else if (numericFormat != null) {
            Element numFmt = elementWithAttr(elementBuilder, "numFmt", "formatCode", numericFormat);
            int count = numFmts.getElementsByTagNameNS(NS_SPREADSHEETML_2006_MAIN, "numFmt").getLength();
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
            fillId = fills.getElementsByTagNameNS(NS_SPREADSHEETML_2006_MAIN, "fill").getLength();

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

            fontId = fonts.getElementsByTagNameNS(NS_SPREADSHEETML_2006_MAIN, "font").getLength() - 1;
        }


        // border handling <border diagonalDown="false" diagonalUp="false"><left/><right/><top/><bottom/><diagonal/></border>
        if (diagonalStyle != null || borderDesc != null) {
            Element border = elementBuilder.apply("border");

            border.setAttribute("diagonalUp", "false");
            border.setAttribute("diagonalDown", "false");

            if (diagonalStyle == DiagonalStyle.BOTTOM_LEFT_TO_TOP_RIGHT) {
                border.setAttribute("diagonalUp", "true");
            } else if (diagonalStyle == DiagonalStyle.TOP_LEFT_TO_BOTTOM_RIGHT) {
                border.setAttribute("diagonalDown", "true");
            } else if (diagonalStyle == DiagonalStyle.BOTH) {
                border.setAttribute("diagonalUp", "true");
                border.setAttribute("diagonalDown", "true");
            }

            Element top = elementBuilder.apply("top");
            Element left = elementBuilder.apply("left");
            Element bottom = elementBuilder.apply("bottom");
            Element right = elementBuilder.apply("right");

            Element diagonal = elementBuilder.apply("diagonal");

            if (diagonalStyle != null) {
                diagonal.setAttribute("style", diagonalLineStyle == null ? LineStyle.THIN.toXmlValue() : diagonalLineStyle.toXmlValue());
                if (diagonalColor != null) {
                    diagonal.appendChild(elementWithAttr(elementBuilder, "color", "rgb", formatColor(diagonalColor)));
                }
            }

            if (borderDesc != null) {
                applyBorder(elementBuilder, top, BorderBuilder.Border.TOP);
                applyBorder(elementBuilder, left, BorderBuilder.Border.LEFT);
                applyBorder(elementBuilder, bottom, BorderBuilder.Border.BOTTOM);
                applyBorder(elementBuilder, right, BorderBuilder.Border.RIGHT);
            }

            border.appendChild(top);
            border.appendChild(left);
            border.appendChild(bottom);
            border.appendChild(right);

            border.appendChild(diagonal);

            borders.appendChild(border);

            borderId = borders.getElementsByTagNameNS(NS_SPREADSHEETML_2006_MAIN, "border").getLength() - 1;
        }

        //

        boolean applyAlignment = readingOrder != null || verticalAlignment != null || horizontalAlignment != null;

        // add the "xf" element
        // <xf numFmtId="164" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false">
        //  <alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/>
        //  <protection locked="true" hidden="false"/>
        // </xf>
        Element xf = elementBuilder.apply("xf");
        xf.setAttribute("numFmtId", Integer.toString(numFmtId));
        xf.setAttribute("fontId", Integer.toString(fontId));
        xf.setAttribute("fillId", Integer.toString(fillId));
        xf.setAttribute("borderId", Integer.toString(borderId));
        xf.setAttribute("xfId", "0");
        xf.setAttribute("applyFont", "true");
        xf.setAttribute("applyBorder", "false");
        xf.setAttribute("applyAlignment", applyAlignment ? "true" : "false");
        xf.setAttribute("applyProtection", "false");
        Element alignment = elementBuilder.apply("alignment");
        alignment.setAttribute("horizontal", horizontalAlignment == null ? "general" : horizontalAlignment.name().toLowerCase(Locale.ROOT));
        alignment.setAttribute("vertical", verticalAlignment == null ? "bottom" : verticalAlignment.name().toLowerCase(Locale.ROOT));
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

        return cellXfs.getElementsByTagNameNS(NS_SPREADSHEETML_2006_MAIN, "xf").getLength() - 1;
    }

    private void applyBorder(Function<String, Element> elementBuilder, Element element, BorderBuilder.Border border) {
        LineStyle lineStyle = borderDesc.style(border);
        String color = borderDesc.color(border);
        if (lineStyle != null || color != null) {
            element.setAttribute("style", lineStyle == null ? LineStyle.THIN.toXmlValue() : lineStyle.toXmlValue());
        }
        if (color != null) {
            element.appendChild(elementWithAttr(elementBuilder, "color", "rgb", formatColor(color)));
        }
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

        //
        private LineStyle diagonalLineStyle;
        private String diagonalColor;
        private DiagonalStyle diagonalStyle;

        private BorderBuilder borderBuilder;
        //
        private VerticalAlignment verticalAlignment;
        private HorizontalAlignment horizontalAlignment;
        //


        StyleBuilder(Function<Style, Boolean> register) {
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
         * Border related options..
         *
         * @return a BorderBuilder
         */
        public BorderBuilder border() {
            if (borderBuilder == null) {
                borderBuilder = new BorderBuilder(this);
            }
            return borderBuilder;
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

        public StyleBuilder verticalAlignment(VerticalAlignment verticalAlignment) {
            this.verticalAlignment = verticalAlignment;
            return this;
        }

        public StyleBuilder horizontalAlignment(HorizontalAlignment horizontalAlignment) {
            this.horizontalAlignment = horizontalAlignment;
            return this;
        }


        //
        public StyleBuilder diagonalStyle(DiagonalStyle diagonalStyle) {
            this.diagonalStyle = diagonalStyle;
            return this;
        }

        public StyleBuilder diagonalLineStyle(LineStyle lineStyle) {
            this.diagonalLineStyle = lineStyle;
            return this;
        }

        public StyleBuilder diagonalColor(String color) {
            this.diagonalColor = color;
            return this;
        }

        public StyleBuilder diagonalColor(Color color) {
            return diagonalColor(color.color);
        }
        //


        /**
         * Generate the style.
         *
         * @return
         */
        public Style build() {
            FontDesc fd = fontBuilder != null ? new FontDesc(fontBuilder.name, fontBuilder.size, fontBuilder.color, fontBuilder.bold,
                    fontBuilder.italic, fontBuilder.fontUnderlineStyle, fontBuilder.strikeOut) : null;

            BorderDesc bd = borderBuilder != null ? new BorderDesc(borderBuilder.color, borderBuilder.style, borderBuilder.borderColor, borderBuilder.borderStyle) : null;

            Style s = new Style(numericFormat, numericFormatIndex, bgColor, fgColor, pattern, rotation,
                    diagonalLineStyle, diagonalColor, diagonalStyle,
                    readingOrder, verticalAlignment, horizontalAlignment, fd, bd);
            register.apply(s);
            return s;
        }
    }

    static class FontDesc {
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

    private static class BorderDesc {
        final String color;
        final LineStyle style;
        final Map<BorderBuilder.Border, String> borderColor;
        final Map<BorderBuilder.Border, LineStyle> borderStyle;

        private BorderDesc(String color, LineStyle style, Map<BorderBuilder.Border, String> borderColor, Map<BorderBuilder.Border, LineStyle> borderStyle) {
            this.color = color;
            this.style = style;
            this.borderColor = borderColor;
            this.borderStyle = borderStyle;
        }

        String color(BorderBuilder.Border border) {
            return borderColor.getOrDefault(border, color);
        }

        LineStyle style(BorderBuilder.Border border) {
            return borderStyle.getOrDefault(border, style);
        }
    }

    /**
     * Border specific options builder.
     */
    public static class BorderBuilder {

        private String color;
        private LineStyle style;
        private Map<Border, String> borderColor = new EnumMap<>(Border.class);
        private Map<Border, LineStyle> borderStyle = new EnumMap<>(Border.class);

        private final StyleBuilder styleBuilder;

        private BorderBuilder(StyleBuilder styleBuilder) {
            this.styleBuilder = styleBuilder;
        }

        public StyleBuilder and() {
            return styleBuilder;
        }

        public Style build() {
            return styleBuilder.build();
        }

        /**
         * Color for all the 4 borders.
         *
         * @param color
         * @return
         */
        public BorderBuilder color(String color) {
            this.color = color;
            return this;
        }

        /**
         * Color for all the 4 borders.
         *
         * @param color
         * @return
         */
        public BorderBuilder color(Color color) {
            return color(color.color);
        }

        /**
         * Line style for all the 4 borders.
         *
         * @param style
         * @return
         */
        public BorderBuilder style(LineStyle style) {
            this.style = style;
            return this;
        }

        /**
         * Define the border color for a specific border.
         *
         * @param border
         * @param color
         * @return
         */
        public BorderBuilder borderColor(Border border, String color) {
            borderColor.put(border, color);
            return this;
        }

        /**
         * Define the border color for a specific border.
         *
         * @param border
         * @param color
         * @return
         */
        public BorderBuilder borderColor(Border border, Color color) {
            return borderColor(border, color.color);
        }

        /**
         * Define the border style for a specific border.
         *
         * @param border
         * @param style
         * @return
         */
        public BorderBuilder borderStyle(Border border, LineStyle style) {
            borderStyle.put(border, style);
            return this;
        }

        /**
         * Which border.
         */
        public enum Border {
            TOP,
            RIGHT,
            BOTTOM,
            LEFT
        }
    }

    static final String DEFAULT_FONT_NAME = "Arial";
    static final BigDecimal DEFAULT_FONT_SIZE = BigDecimal.TEN;

    /**
     * Font specific options.
     */
    public static class FontBuilder {
        String name = DEFAULT_FONT_NAME;
        BigDecimal size = DEFAULT_FONT_SIZE;
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

        /**
         * Bold variant.
         *
         * @param bold
         * @return
         */
        public FontBuilder bold(boolean bold) {
            this.bold = bold;
            return this;
        }

        /**
         * Italic variant.
         *
         * @param italic
         * @return
         */
        public FontBuilder italic(boolean italic) {
            this.italic = italic;
            return this;
        }

        /**
         * Name of the font (e.g.: Arial, Verdana, ...)
         *
         * @param name
         * @return
         */
        public FontBuilder name(String name) {
            Objects.requireNonNull(name);
            this.name = name;
            return this;
        }

        /**
         * Define the underline style.
         *
         * @param fontUnderlineStyle
         * @return
         */
        public FontBuilder underline(FontUnderlineStyle fontUnderlineStyle) {
            Objects.requireNonNull(fontUnderlineStyle);
            this.fontUnderlineStyle = fontUnderlineStyle;
            return this;
        }

        /**
         * Strike the text.
         *
         * @param strikeOut
         * @return
         */
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

        /**
         * Set the font with a given color.
         *
         * @param color
         * @return
         */
        public FontBuilder color(Color color) {
            return color(color.color);
        }

        /**
         * Font size.
         *
         * @param size
         * @return
         */
        public FontBuilder size(int size) {
            this.size = BigDecimal.valueOf(size);
            return this;
        }

        /**
         * Font size.
         *
         * @param size
         * @return
         */
        public FontBuilder size(double size) {
            this.size = BigDecimal.valueOf(size);
            return this;
        }

        /**
         * Font size.
         *
         * @param size
         * @return
         */
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


    /**
     * Pattern as defined in https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.patternvalues?view=openxml-2.8.1
     */
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

    /**
     * Line/Border style as defined in https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.borderstylevalues?view=openxml-2.8.1
     */
    public enum LineStyle {
        DASH_DOT("dashDot"),
        DASH_DOT_DOT("dashDotDot"),
        DASHED,
        DOTTED,
        DOUBLE,
        HAIR,
        MEDIUM,
        MEDIUM_DASH_DOT("mediumDashDot"),
        MEDIUM_DASH_DOT_DOT("mediumDashDotDot"),
        MEDIUM_DASHED("mediumDashed"),
        NONE,
        SLANT_DASH_DOT("slantDashDot"),
        THICK,
        THIN;

        private final String xmlName;

        LineStyle(String xmlName) {
            this.xmlName = xmlName;
        }

        LineStyle() {
            this(null);
        }

        String toXmlValue() {
            return xmlName == null ? name().toLowerCase(Locale.ROOT) : xmlName;
        }
    }

    public enum DiagonalStyle {
        BOTTOM_LEFT_TO_TOP_RIGHT,
        TOP_LEFT_TO_BOTTOM_RIGHT,
        BOTH;
    }

    public enum HorizontalAlignment {
        LEFT,
        CENTER,
        RIGHT,
        DISTRIBUTED,
        FILL,
        GENERAL,
        JUSTIFY
    }

    public enum VerticalAlignment {
        TOP,
        CENTER,
        BOTTOM,
        DISTRIBUTED,
        JUSTIFY
    }
}
