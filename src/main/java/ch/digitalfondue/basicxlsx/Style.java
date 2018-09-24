package ch.digitalfondue.basicxlsx;

//based on https://xlsxwriter.readthedocs.io/format.html
public class Style {

    private Style() {
    }

    public static class StyleBuilder {

        private FontBuilder fontBuilder;

        private StyleBuilder() {
        }

        public FontBuilder font() {
            if (fontBuilder == null) {
                fontBuilder = new FontBuilder(this);
            }
            return fontBuilder;
        }

        public Style build() {
            return new Style();
        }
    }

    public static class FontBuilder {
        String name;
        Integer size;
        String color;
        Boolean bold;
        Boolean italic;
        FontUnderlineStyle fontUnderlineStyle;
        Boolean strikeOut;

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


    public static StyleBuilder define() {
        return new StyleBuilder();
    }
}
