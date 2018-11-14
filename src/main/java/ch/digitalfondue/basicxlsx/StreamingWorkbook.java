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

import org.w3c.dom.Element;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * <p>Represent a xlsx workbook. It's the main entry point for generating a xlsx file.</p>
 * <p>This workbook, as his name implies, allow to stream the rows one by one, thus avoiding to keep all
 * the cells data in memory. This mode has some downsides (no auto sizing for column), but it's useful if
 * you are memory constrained and/or you don't need all the features of the more complete {@link Workbook}.</p>
 *
 * <p>Note: this class implements {@link AutoCloseable} and it's best used within a try-with-resources statement.
 * If it's not in a try-with-resources statement, then you _must_ call {@link #close()} or else the
 * xlsx file will be incomplete.</p>
 */
public class StreamingWorkbook extends AbstractWorkbook implements Closeable, AutoCloseable {

    /**
     * Row wrapper.
     * */
    public static class Row {

        private final Cell[] cells;
        private final Double height;

        Row(Cell[] cells, Double height) {
            this.cells = cells;
            this.height = height;
        }
    }

    /**
     * Configuration options for a given sheet.
     */
    public static class SheetOptions {
        private double[] columnWidth;
        private Style.ReadingOrder readingOrder;

        public SheetOptions(double[] columnWidth) {
            this(columnWidth, null);
        }

        public SheetOptions(Style.ReadingOrder readingOrder) {
            this(null, readingOrder);
        }

        public SheetOptions(double[] columnWidth, Style.ReadingOrder readingOrder) {
            this.columnWidth = columnWidth;
            this.readingOrder = readingOrder;
        }
    }

    /**
     * Wrap an array of cells in a {@link Row}
     *
     * @param cells
     * @return
     */
    public static Row row(Cell[] cells) {
        return new Row(cells, null);
    }

    /**
     * Wrap a collection of cells in a {@link Row}
     *
     * @param cells
     * @return
     */
    public static Row row(Collection<Cell> cells) {
        return row(cells.toArray(new Cell[cells.size()]));
    }

    /**
     * Wrap an array of cells in a {@link Row}, with a specific height.
     *
     * @param cells
     * @return
     */
    public static Row row(Cell[] cells, double height) {
        return new Row(cells, height);
    }

    /**
     * Wrap a collection of cells in a {@link Row}, with a specific height.
     *
     * @param cells
     * @param height
     * @return
     */
    public static Row row(Collection<Cell> cells, double height) {
        return row(cells.toArray(new Cell[cells.size()]), height);
    }

    private final ZipOutputStream zos;
    private boolean hasEnded;
    private boolean hasRegisteredStyles;
    private List<String> sheets = new ArrayList<>();

    private static final byte[] SHEET_START = ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n").getBytes(StandardCharsets.UTF_8);

    private static final byte[] DEFAULT_COL = "<col max=\"1\" min=\"1\"/>".getBytes(StandardCharsets.UTF_8);
    private static final byte[] SHEET_END_COLS = "</cols><sheetData>".getBytes(StandardCharsets.UTF_8);

    private static final byte[] SHEET_END = "</sheetData></worksheet>".getBytes(StandardCharsets.UTF_8);

    private static final byte[] ROW_START_1 = "<row r=\"".getBytes(StandardCharsets.UTF_8);
    private static final byte[] ROW_START_2 = "\">".getBytes(StandardCharsets.UTF_8);
    private static final byte[] ROW_END = "</row>".getBytes(StandardCharsets.UTF_8);
    private final Function<String, Element> elementBuilder;

    public StreamingWorkbook(OutputStream os) {
        this.zos = new ZipOutputStream(os, StandardCharsets.UTF_8);
        this.elementBuilder = Utils.toElementBuilder(Utils.toDocument("ch/digitalfondue/basicxlsx/sheet_template.xml"));
    }

    /**
     * Close the workbook.
     *
     * @throws IOException
     */
    @Override
    public void close() throws IOException {
        if (!hasEnded) {
            end();
        }
        zos.close();
    }

    /**
     * Define a new style.
     *
     * @return a style builder
     */
    @Override
    public Style.StyleBuilder defineStyle() {
        if (hasRegisteredStyles) {
            throw new IllegalStateException("Cannot register new styles after writing a sheet");
        }
        return super.defineStyle();
    }

    /**
     * Write a sheet.
     *
     * @param name: the name of the sheet
     * @param rows: the rows
     * @throws IOException
     */
    public void withSheet(String name, Stream<Row> rows) throws IOException {
        withSheet(name, rows, null);
    }


    private void write(String s) throws IOException {
        zos.write(s.getBytes(StandardCharsets.UTF_8));
    }

    /**
     * Write a sheet (with some options).
     *
     * @param name
     * @param rows
     * @param options
     * @throws IOException
     */
    public void withSheet(String name, Stream<Row> rows, SheetOptions options) throws IOException {
        if (hasEnded) {
            throw new IllegalStateException("Already ended");
        }

        if (!hasRegisteredStyles) {
            commitAndWriteStyleMetadata(zos, styles, styleToIdMapping);
            hasRegisteredStyles = true;
        }

        sheets.add(name);

        zos.putNextEntry(new ZipEntry("xl/worksheets/sheet" + (sheets.size()) + ".xml"));
        zos.write(SHEET_START);

        if (options != null && options.readingOrder != null) {
            write("<sheetViews><sheetView rightToLeft=\"");
            write(Boolean.toString(options.readingOrder == Style.ReadingOrder.RTL));
            write("\"></sheetView></sheetViews>");
        }

        write("<cols>");
        if (options == null || options.columnWidth == null || options.columnWidth.length == 0) {
            zos.write(DEFAULT_COL);
        } else {
            for (int i = 0; i < options.columnWidth.length; i++) {
                double colWidth = options.columnWidth[i];
                writeCol(i, colWidth);
            }
        }

        zos.write(SHEET_END_COLS);
        AtomicInteger rowCounter = new AtomicInteger(0);

        Transformer transformer = Utils.getTransformer(true);
        StreamResult sr = new StreamResult(new OutputStreamWriter(zos, StandardCharsets.UTF_8));
        Consumer<DOMSource> consumer = domSource -> {
            try {
                transformer.transform(domSource, sr);
            } catch (TransformerException e) {
                throw new IllegalStateException(e);
            }
        };

        rows.forEachOrdered(row -> {
            processRow(rowCounter.get(), row, consumer);
            rowCounter.incrementAndGet(); //ugly, but it works
        });
        zos.write(SHEET_END);
        zos.closeEntry();
    }

    private void writeCol(int idx, double colWidth) throws IOException {
        byte[] minMax = Integer.toString(idx + 1).getBytes(StandardCharsets.UTF_8);
        write("<col max=\"");
        zos.write(minMax);
        write("\" min=\"");
        zos.write(minMax);

        if (colWidth > 0) {
            write("\" customWidth=\"true\" width=\"");
            write(Double.toString(colWidth));
        }

        write("\"/>");
    }

    private void processRow(int rowIdx, Row rowContainer, Consumer<DOMSource> consumer) {
        try {
            if (rowContainer != null && rowContainer.cells != null) {
                Cell[] row = rowContainer.cells;
                //"<row r="1">"
                zos.write(ROW_START_1);
                write(Integer.toString(rowIdx + 1));

                if (rowContainer.height != null) {
                    write("\" customHeight=\"true\" ht=\"");
                    write(Double.toString(rowContainer.height));
                }
                zos.write(ROW_START_2);
                //
                for (int i = 0; i < row.length; i++) {
                    Cell cell = row[i];
                    if (cell != null) {
                        int styleId = styleIdSupplier(cell);
                        Element e = cell.toElement(elementBuilder, rowIdx, i, styleId);
                        //TODO: find a way to remove the xmlns attached to the cell...
                        consumer.accept(new DOMSource(e));
                    }
                }
                zos.write(ROW_END);
            }
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    private void end() throws IOException {
        if (hasEnded) {
            throw new IllegalStateException("already ended");
        } else {
            hasEnded = true;
            writeMetadataDocuments(zos, sheets);
        }
    }
}
