package ch.digitalfondue.basicxlsx;

import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

public class WorkbookTest {

    @Test
    public void testWorkbookCreation() throws IOException {
        Workbook w = new Workbook();

        Sheet s = w.sheet("test");
        s.setValueAt("Hello", 0, 0); //A1
        s.setValueAt("B1", 0, 1);

        s.setValueAt("A2", 1, 0);
        s.setValueAt("World!", 1, 1); //B2

        Sheet s2 = w.sheet("test2");
        s2.setValueAt("Hello", 1, 0); //A2
        s2.setValueAt("World", 0, 1); //B1

        try (FileOutputStream fos = new FileOutputStream("test.xlsx")) {
            w.write(fos);
        }
    }
}
