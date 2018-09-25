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

import org.junit.Assert;
import org.junit.Test;

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
}
