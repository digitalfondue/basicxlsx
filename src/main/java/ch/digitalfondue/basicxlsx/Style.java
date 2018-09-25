package ch.digitalfondue.basicxlsx;

import org.w3c.dom.Element;

import java.util.function.Function;

//based on https://xlsxwriter.readthedocs.io/format.html
public class Style {

    private Style() {
    }

    void register(Element fonts, Element cellXfs) {
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
            Style s =  new Style();
            register.apply(s);
            return s;
        }
    }

    public static class FontBuilder {
        String name = "Arial";
        int size = 10;//default
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
            return this.styleBuilder;
        }

        public FontBuilder bold(boolean bold) {
            this.bold = bold;
            return this;
        }

        public FontBuilder italic(boolean italic) {
            this.italic = italic;
            return this;
        }
    }

    public enum FontUnderlineStyle {
        NONE, SINGLE, DOUBLE, SINGLE_ACCOUNTING_UNDERLINE, DOUBLE_ACCOUNTING_UNDERLINE;
    }


    static StyleBuilder define(Function<Style, Boolean> register) {
        return new StyleBuilder(register);
    }
}
