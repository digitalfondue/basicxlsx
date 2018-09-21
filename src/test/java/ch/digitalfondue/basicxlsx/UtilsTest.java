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
