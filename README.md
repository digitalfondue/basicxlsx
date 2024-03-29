# basicxlsx: when you only need to write simple xlsx

[![Maven Central](https://img.shields.io/maven-central/v/ch.digitalfondue.basicxlsx/basicxlsx.svg)](https://search.maven.org/search?q=g:ch.digitalfondue.basicxlsx)
[![Build Status](https://img.shields.io/github/actions/workflow/status/digitalfondue/basicxslx/.github/workflows/maven.yml)](https://github.com/digitalfondue/basicxlsx/actions?query=workflow%3A%22Java+CI+with+Maven%22)

## Why

Apache POI, being a complete solution, is quite heavyweight.
This library provide only the minimal amount of functionality to
write xlsx files.

## License
basicxlsx is licensed under the Apache License Version 2.0.

## Download

maven:

```xml
<dependency>
    <groupId>ch.digitalfondue.basicxlsx</groupId>
    <artifactId>basicxlsx</artifactId>
    <version>0.7.1</version>
</dependency>
```

gradle:

```
compile 'ch.digitalfondue.basicxlsx:basicxlsx:0.7.1'
```

## Javadoc

Available at [javadoc.io](http://javadoc.io/doc/ch.digitalfondue.basicxlsx/basicxlsx/).
 
The api is minimalistic, basically you only need to choose Workbook type (streaming or not),
 define your styles, define the data and write the workbook to an OutputStream. 

If you use it as a module, remember to add requires `ch.digitalfondue.basicxlsx;` in your module-info.

## Example

### Minimal example

```java
import ch.digitalfondue.basicxlsx.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class Example {

    public static void main(String[] args) throws IOException {
        Workbook w = new Workbook();
        Sheet s = w.sheet("test"); //create a new sheet named test
        s.setValueAt("Hello World", /*row*/ 0, /*column*/ 0); //put in "A1" the value "Hello World"

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

    public static void main(String[] args) throws IOException {
        Workbook w = new Workbook();


        // you must define the styles before using them
        Style redBGBold = w.defineStyle().bgColor(Style.Color.RED).font().bold(true).build();
        //

        Sheet s = w.sheet("test"); //create a new sheet named test
        s.setValueAt("Hello World", /*row*/ 0, /*column*/ 0).withStyle(redBGBold); //put in "A1" the value "Hello World", set the style to the cell

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
import ch.digitalfondue.basicxlsx.StreamingWorkbook.Row;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

public class ExampleStreaming {

    public static void main(String[] args) throws IOException {

        try (FileOutputStream fos = new FileOutputStream("test.xlsx");
             StreamingWorkbook w = new StreamingWorkbook(fos)) { //<- create a StreamingWorkbook: it require the outputstream


            Cell[] row1 = new Cell[] {Cell.cell("Hello World")};
            Stream<Row> rows = Stream.of(StreamingWorkbook.row(row1));

            w.withSheet("test", rows); //write a new sheet named "test" with the stream of rows
        }
    }
}
```


### Streaming example with style

```java
import ch.digitalfondue.basicxlsx.Cell;
import ch.digitalfondue.basicxlsx.StreamingWorkbook;
import ch.digitalfondue.basicxlsx.StreamingWorkbook.Row;
import ch.digitalfondue.basicxlsx.Style;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

public class ExampleStreamingWithStyle {

    public static void main(String[] args) throws IOException {

        try (FileOutputStream fos = new FileOutputStream("test.xlsx");
             StreamingWorkbook w = new StreamingWorkbook(fos)) { //<- create a StreamingWorkbook: it require the outputstream

            // you must define the styles before
            Style redBGBold = w.defineStyle().bgColor(Style.Color.RED).font().bold(true).build();
            //

            Cell[] row1 = new Cell[] {Cell.cell("Hello World").withStyle(redBGBold)};
            Stream<Row> rows = Stream.of(StreamingWorkbook.row(row1));


            w.withSheet("test", rows); //write a new sheet named "test" with the stream of rows
        }
    }
}
```

See https://github.com/digitalfondue/basicxlsx/blob/master/src/test/java/ch/digitalfondue/basicxlsx/StreamingWorkbookTest.java
for a more complete example with style, formatting and other data types.


## TODO:

 - [ ] add more test
 - [ ] write javadoc
 - [ ] merged cell
 - [ ] charts
 - [ ] conditional formatting
 - [ ] ... ?
 - styling
    - [ ] complete missing alignment
    - [ ] indent https://xlsxwriter.readthedocs.io/format.html#format-set-indent (some gotcha are present it seems...)
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
