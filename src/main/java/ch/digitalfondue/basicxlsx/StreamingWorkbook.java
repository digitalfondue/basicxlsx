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
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class StreamingWorkbook extends AbstractWorkbook implements Closeable, AutoCloseable {

    public static class Row {

        private final Cell[] cells;

        Row(Cell[] cells) {
            this.cells = cells;
        }
    }

    public static Row row(Cell[] cells) {
        return new Row(cells);
    }

    private final ZipOutputStream zos;
    private boolean hasEnded;
    private boolean hasRegisteredStyles;
    private List<String> sheets = new ArrayList<>();

    private static final byte[] SHEET_START = ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" +
            "<cols>").getBytes(StandardCharsets.UTF_8);

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

    @Override
    public void close() throws IOException {
        if (!hasEnded) {
            end();
        }
        zos.close();
    }

    @Override
    public Style.StyleBuilder defineStyle() {
        if (hasRegisteredStyles) {
            throw new IllegalStateException("Cannot register new styles after writing a sheet");
        }
        return super.defineStyle();
    }

    public void withSheet(String name, Stream<Row> rows) throws IOException {
        withSheet(name, rows, null);
    }

    public void withSheet(String name, Stream<Row> rows, double[] columnWidth) throws IOException {
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
        if (columnWidth == null || columnWidth.length == 0) {
            zos.write(DEFAULT_COL);
        } else {
            for (int i = 0; i < columnWidth.length; i++) {
                double colWidth = columnWidth[i];
                writeCol(i, colWidth, zos);
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
            processRow(rowCounter.get(), row.cells, consumer);
            rowCounter.incrementAndGet(); //ugly, but it works
        });
        zos.write(SHEET_END);
        zos.closeEntry();
    }

    private static void writeCol(int idx, double colWidth, ZipOutputStream zos) throws IOException {
        byte[] minMax = Integer.toString(idx + 1).getBytes(StandardCharsets.UTF_8);
        zos.write("<col max=\"".getBytes(StandardCharsets.UTF_8));
        zos.write(minMax);
        zos.write("\" min=\"".getBytes(StandardCharsets.UTF_8));
        zos.write(minMax);

        if (colWidth > 0) {
            zos.write("\" customWidth=\"true\" width=\"".getBytes(StandardCharsets.UTF_8));
            zos.write(Double.toString(colWidth).getBytes(StandardCharsets.UTF_8));
        }

        zos.write("\"/>".getBytes(StandardCharsets.UTF_8));
    }

    private void processRow(int rowIdx, Cell[] row, Consumer<DOMSource> consumer) {
        try {
            if (row != null) {
                //"<row r="1">"
                zos.write(ROW_START_1);
                zos.write(Integer.toString(rowIdx + 1).getBytes(StandardCharsets.UTF_8));
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

    public void end() throws IOException {
        if (hasEnded) {
            throw new IllegalStateException("already ended");
        } else {
            hasEnded = true;
            writeMetadataDocuments(zos, sheets);
        }
    }
}
