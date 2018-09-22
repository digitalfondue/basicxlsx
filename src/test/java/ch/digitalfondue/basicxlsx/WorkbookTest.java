package ch.digitalfondue.basicxlsx;

import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;

public class WorkbookTest {

    @Test
    public void testWorkbookCreation() throws IOException {
        Workbook w = new Workbook();

        Sheet s = w.sheet("test");
        s.setValueAt("Hello éé èè Michał", 0, 0); //A1
        s.setValueAt("B1", 0, 1);

        s.setValueAt("A2", 1, 0);
        s.setValueAt("World!", 1, 1); //B2

        s.setValueAt(42, 2,2); //C3
        s.setValueAt(new BigDecimal("2.512351234324832"), 3,2); //C4
        s.setValueAt(3.14, 4,2); //C5

        Sheet s2 = w.sheet("test2");
        s2.setValueAt("Hello", 1, 0); //A2
        s2.setValueAt("World", 0, 1); //B1

        try (FileOutputStream fos = new FileOutputStream("test.xlsx")) {
            w.write(fos);
        }
    }
}
