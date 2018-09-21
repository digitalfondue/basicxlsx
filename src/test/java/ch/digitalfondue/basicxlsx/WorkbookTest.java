package ch.digitalfondue.basicxlsx;

import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

public class WorkbookTest {

    @Test
    public void testWorkbookCreation() throws IOException {
        Workbook w = new Workbook();

        Sheet s = w.openSheet("test");
        s.setCellAt(new Cell("Hello"), 0, 0); //A1
        s.setCellAt(new Cell("B1"), 0, 1);

        s.setCellAt(new Cell("A2"), 1, 0);
        s.setCellAt(new Cell("World!"), 1,1); //B2

        Sheet s2 = w.openSheet("test2");
        s2.setCellAt(new Cell("Hello"), 1, 0); //A2
        s2.setCellAt(new Cell("World"), 0, 1); //B1

        try (FileOutputStream fos = new FileOutputStream("test.xlsx")) {
            w.write(fos);
        }
    }
}
