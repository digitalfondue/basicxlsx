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

import java.awt.font.FontRenderContext;
import java.awt.font.TextAttribute;
import java.awt.font.TextLayout;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.math.BigDecimal;
import java.text.AttributedString;

class CellWidthCalculator {

    // https://support.microsoft.com/en-ph/help/214123/description-of-how-column-widths-are-determined-in-excel
    // currently only handle string and boolean type, as we don't support the formatting
    static double cellWidth(Cell cell) {

        String value = cell.formattedValue();
        if (value == null) {
            return 8.43; //standard size
        }

        if (cell.getStyle() == null) {
            return value.length() * 0.9;
        } else {
            return getWidth(value, cell.getStyle());
        }
    }
    //

    // code imported and modified from https://github.com/apache/poi/blob/trunk/src/java/org/apache/poi/ss/util/SheetUtil.java
    // under the following license
    /* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at
       http://www.apache.org/licenses/LICENSE-2.0
   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
   ==================================================================== */
    //

    private static final FontRenderContext FONT_RENDER_CONTEXT = new FontRenderContext(null, true, true);

    private static float getDefaultCharWidth() {
        AttributedString str = new AttributedString(String.valueOf('0'));
        copyAttributes(new Style.FontDesc("Arial", BigDecimal.TEN, null, false, false, null, false), str, 0, 1);
        TextLayout layout = new TextLayout(str.getIterator(), FONT_RENDER_CONTEXT);
        return layout.getAdvance();
    }

    private static double getWidth(String value, Style style) {
        AttributedString str = new AttributedString(value);
        copyAttributes(style.getFontDesc(), str, 0, value.length());
        TextLayout layout = new TextLayout(str.getIterator(), FONT_RENDER_CONTEXT);
        Rectangle2D bounds;
        if(style.getRotation() != null && style.getRotation() != 0){
            /*
             * Transform the text using a scale so that it's height is increased by a multiple of the leading,
             * and then rotate the text before computing the bounds. The scale results in some whitespace around
             * the unrotated top and bottom of the text that normally wouldn't be present if unscaled, but
             * is added by the standard Excel autosize.
             */
            AffineTransform trans = new AffineTransform();
            trans.concatenate(AffineTransform.getRotateInstance(style.getRotation() * 2.0 * Math.PI / 360.0));
            trans.concatenate(AffineTransform.getScaleInstance(1, 2.0));
            bounds = layout.getOutline(trans).getBounds();
        } else {
            bounds = layout.getBounds();
        }
        // frameWidth accounts for leading spaces which is excluded from bounds.getWidth()
        double frameWidth = bounds.getX() + bounds.getWidth();
        return ((frameWidth) / getDefaultCharWidth());
    }

    private static void copyAttributes(Style.FontDesc font, AttributedString str, int startIdx, int endIdx) {
        str.addAttribute(TextAttribute.FAMILY, font.name, startIdx, endIdx);
        str.addAttribute(TextAttribute.SIZE, font.size.floatValue());
        if (font.bold) {
            str.addAttribute(TextAttribute.WEIGHT, TextAttribute.WEIGHT_BOLD, startIdx, endIdx);
        }
        if (font.italic) {
            str.addAttribute(TextAttribute.POSTURE, TextAttribute.POSTURE_OBLIQUE, startIdx, endIdx);
        }
        if (font.fontUnderlineStyle != Style.FontUnderlineStyle.NONE && font.fontUnderlineStyle != null) {
            str.addAttribute(TextAttribute.UNDERLINE, TextAttribute.UNDERLINE_ON, startIdx, endIdx);
        }
    }
}
