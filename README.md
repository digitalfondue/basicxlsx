# basicxlsx: when you only need to write simple xlsx

[![Maven Central](https://img.shields.io/maven-central/v/ch.digitalfondue.basicxlsx/basicxlsx.svg)](https://search.maven.org/search?q=g:ch.digitalfondue.basicxlsx)
[![Build Status](https://travis-ci.org/digitalfondue/basicxlsx.svg?branch=master)](https://travis-ci.org/digitalfondue/basicxlsx) 
[![Coverage Status](https://coveralls.io/repos/github/digitalfondue/basicxlsx/badge.svg?branch=master)](https://coveralls.io/github/digitalfondue/basicxlsx?branch=master)

## Why

Apache POI, being a complete solution, is quite heavyweight.
This library provide only the minimal amount of functionality to
write xslx files.

## License
basicxlsx is licensed under the Apache License Version 2.0.

## Download

maven:

```xml
<dependency>
    <groupId>ch.digitalfondue.basicxlsx</groupId>
    <artifactId>basicxlsx</artifactId>
    <version>0.2.0</version>
</dependency>
```

gradle:

```
compile 'ch.digitalfondue.basicxlsx:basicxlsx:0.2.0'
```

## Example

### Minimal example

```java
import ch.digitalfondue.basicxlsx.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class Example {

    public static void main(String args[]) throws IOException {
        Workbook w = new Workbook();
        Sheet s = w.sheet("test"); //create a new sheet named test
        s.setValueAt("Hello World", 0, 0); //put in "A1" the value "Hello World"

        //write the workbook to a file
        try (FileOutputStream fos = new FileOutputStream("test.xlsx")) {
            w.write(fos);
        }
    }
}

```

### Style

```java
import ch.digitalfondue.basicxlsx.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExampleWithStyle {

    public static void main(String args[]) throws IOException {
        Workbook w = new Workbook();


        // you must define the styles before using them
        Style redBGBold = w.defineStyle().bgColor(Style.Color.RED).font().bold(true).build();
        //

        Sheet s = w.sheet("test"); //create a new sheet named test
        s.setValueAt("Hello World", 0, 0).withStyle(redBGBold); //put in "A1" the value "Hello World", set the style to the cell

        //write the workbook to a file
        try (FileOutputStream fos = new FileOutputStream("test.xlsx")) {
            w.write(fos);
        }
    }
}

```

See https://github.com/digitalfondue/basicxlsx/blob/master/src/test/java/ch/digitalfondue/basicxlsx/WorkbookTest.java
for a more complete example with style, formatting and other data types.

### Minimal streaming example

```java
import ch.digitalfondue.basicxlsx.Cell;
import ch.digitalfondue.basicxlsx.StreamingWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

public class ExampleStreaming {

    public static void main(String args[]) throws IOException {

        try (FileOutputStream fos = new FileOutputStream("test.xlsx");
             StreamingWorkbook w = new StreamingWorkbook(fos)) { //<- create a StreamingWorkbook: it require the outputstream


            Cell[] row1 = new Cell[] {Cell.cell("Hello World")};
            Stream<Cell[]> rows = Stream.<Cell[]>of(row1);

            w.withSheet("test", rows); //write a new sheet named "test" with the stream of rows
        }
    }
}
```


### Streaming example with style

```java
import ch.digitalfondue.basicxlsx.Cell;
import ch.digitalfondue.basicxlsx.StreamingWorkbook;
import ch.digitalfondue.basicxlsx.Style;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

public class ExampleStreamingWithStyle {

    public static void main(String args[]) throws IOException {

        try (FileOutputStream fos = new FileOutputStream("test.xlsx");
             StreamingWorkbook w = new StreamingWorkbook(fos)) { //<- create a StreamingWorkbook: it require the outputstream

            // you must define the styles before
            Style redBGBold = w.defineStyle().bgColor(Style.Color.RED).font().bold(true).build();
            //

            Cell[] row1 = new Cell[] {Cell.cell("Hello World").withStyle(redBGBold)};
            Stream<Cell[]> rows = Stream.<Cell[]>of(row1);


            w.withSheet("test", rows); //write a new sheet named "test" with the stream of rows
        }
    }
}
```

See https://github.com/digitalfondue/basicxlsx/blob/master/src/test/java/ch/digitalfondue/basicxlsx/StreamingWorkbookTest.java
for a more complete example with style, formatting and other data types.


## Javadoc

Available at [javadoc.io](http://javadoc.io/doc/ch.digitalfondue.basicxlsx/basicxlsx/)


## TODO:

 - support autoSizeColumn https://poi.apache.org/apidocs/org/apache/poi/xssf/usermodel/XSSFSheet.html#autoSizeColumn-int-
    - https://github.com/apache/poi/blob/trunk/src/ooxml/java/org/apache/poi/xssf/streaming/SXSSFSheet.java#L1571
    - https://github.com/apache/poi/blob/trunk/src/java/org/apache/poi/ss/util/SheetUtil.java#L120
    - https://github.com/dtjohnson/xlsx-populate/issues/26#issuecomment-288796920
    - https://metacpan.org/pod/Spreadsheet::WriteExcel::Examples#Example:-autofit.pl
 - support other column type
    - [ ] formula, note as described in https://xlsxwriter.readthedocs.io/working_with_formulas.html#formula-results ,
          we can let excel recalculate all the formula result: WIP, if we write "0" as a placeholder value, 
          libreoffice will use it and ignore the recalculate parameter, need to test if leaving it blank work in excel&co?
    - [ ] missing date type (Zoned* variant?)
 - add test (WIP)
 - [ ] improve column sizing in streaming
 - [ ] write javadoc
 - [ ] merged cell
 - [ ] charts
 - [ ] conditional formatting
 - [ ] ... ?
 - styling
    - [ ] alignment 
    - [ ] cell border
    - [ ] diagonal border
    - [ ] rotation *easy to implement* https://xlsxwriter.readthedocs.io/format.html#format-set-rotation
    - [ ] reading order (RTL/LTR) *easy to implement* https://xlsxwriter.readthedocs.io/format.html#format-set-reading-order
    - [ ] indent *easy to implement* https://xlsxwriter.readthedocs.io/format.html#format-set-indent
    - see for other missing https://xlsxwriter.readthedocs.io/format.html

## Resources/examples about xlsx

- http://officeopenxml.com/anatomyofOOXML-xlsx.php
- https://social.technet.microsoft.com/wiki/contents/articles/19601.powershell-generate-real-excel-xlsx-files-without-excel.aspx
- https://github.com/mk-j/PHP_XLSXWriter
- https://github.com/jmcnamara/XlsxWriter

## Notes

### License format
- `mvn com.mycila:license-maven-plugin:format`

### Check updates
- `mvn versions:display-dependency-updates`
- `mvn versions:display-plugin-updates`
