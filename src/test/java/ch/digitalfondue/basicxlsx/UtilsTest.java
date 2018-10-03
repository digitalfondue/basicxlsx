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

import org.apache.commons.lang3.time.DateUtils;
import org.junit.Assert;
import org.junit.Test;

import java.math.BigDecimal;
import java.text.ParseException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class UtilsTest {

    @Test
    public void fromRowColumnToExcelTest() {
        //A3
        Assert.assertEquals("A3", Utils.fromRowColumnToExcelCoordinates(2,0));
        //Z3
        Assert.assertEquals("Z3", Utils.fromRowColumnToExcelCoordinates(2,25));
        //AA3
        Assert.assertEquals("AA3", Utils.fromRowColumnToExcelCoordinates(2,26));
        //ZZ3
        Assert.assertEquals("ZZ3", Utils.fromRowColumnToExcelCoordinates(2,701));
        //AAA3
        Assert.assertEquals("AAA3", Utils.fromRowColumnToExcelCoordinates(2,702));
    }

    //test data imported from https://github.com/dtjohnson/xlsx-populate/blob/master/test/unit/dateConverter.spec.js
    @Test
    public void dateTest() throws ParseException {
        Assert.assertEquals(BigDecimal.valueOf(1), Utils.getExcelDate(DateUtils.parseDateStrictly("1900-01-01 00:00:00", "yyyy-MM-dd HH:mm:ss")));
        Assert.assertEquals(BigDecimal.valueOf(59), Utils.getExcelDate(DateUtils.parseDateStrictly("1900-02-28 00:00:00", "yyyy-MM-dd HH:mm:ss")));
        Assert.assertEquals(BigDecimal.valueOf(61), Utils.getExcelDate(DateUtils.parseDateStrictly("1900-03-01 00:00:00", "yyyy-MM-dd HH:mm:ss")));
        Assert.assertEquals(new BigDecimal("42070.5599999999976716935634613037109375"), Utils.getExcelDate(DateUtils.parseDateStrictly("2015-03-07 13:26:24", "yyyy-MM-dd HH:mm:ss")));
        Assert.assertEquals(new BigDecimal("42829.8333333333357586525380611419677734375"), Utils.getExcelDate(DateUtils.parseDateStrictly("2017-04-04 20:00:00", "yyyy-MM-dd HH:mm:ss")));
    }

    @Test
    public void localDateTimeTest() {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        Assert.assertEquals(BigDecimal.valueOf(1), Utils.getExcelDate(LocalDateTime.parse("1900-01-01 00:00:00", formatter)));
        Assert.assertEquals(BigDecimal.valueOf(59), Utils.getExcelDate(LocalDateTime.parse("1900-02-28 00:00:00", formatter)));
        Assert.assertEquals(BigDecimal.valueOf(61), Utils.getExcelDate(LocalDateTime.parse("1900-03-01 00:00:00", formatter)));
        Assert.assertEquals(new BigDecimal("42070.5599999999976716935634613037109375"), Utils.getExcelDate(LocalDateTime.parse("2015-03-07 13:26:24", formatter)));
        Assert.assertEquals(new BigDecimal("42829.8333333333357586525380611419677734375"), Utils.getExcelDate(LocalDateTime.parse("2017-04-04 20:00:00", formatter)));
    }
}
