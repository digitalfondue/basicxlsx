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

import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class StreamingWorkbook extends AbstractWorkbook implements Closeable, AutoCloseable {

    private final ZipOutputStream zos;
    private boolean hasEnded;
    private List<String> sheets = new ArrayList<>();

    private static final byte[] SHEET_START = ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" +
            "<cols></cols><sheetData>").getBytes(StandardCharsets.UTF_8);

    private static final byte[] SHEET_END = "</sheetData></worksheet>".getBytes(StandardCharsets.UTF_8);
    private static final byte[] ROW_END = "</row>".getBytes(StandardCharsets.UTF_8);

    public StreamingWorkbook(OutputStream os) {
        this.zos = new ZipOutputStream(os, StandardCharsets.UTF_8);
    }

    @Override
    public void close() throws IOException {
        if (!hasEnded) {
            end();
        }
        zos.close();
    }

    public void withSheet(String name, Stream<Cell[]> rows) throws IOException {
        sheets.add(name);

        zos.putNextEntry(new ZipEntry("xl/worksheets/sheet" + (sheets.size()) + ".xml"));
        zos.write(SHEET_START);
        AtomicInteger rowCounter = new AtomicInteger(0);
        rows.forEachOrdered(row -> {
            processRow(rowCounter.get(), row);
            rowCounter.incrementAndGet(); //ugly, but it works
        });
        zos.write(SHEET_END);
        zos.closeEntry();
    }

    private void processRow(int rowIdx, Cell[] row) {
        try {
            //zos.write("<row r="1">"); FIXME implement

            if (row != null) {
                for (int i = 0; i < row.length; i++) {
                    Cell cell = row[i];
                    if (cell != null) {
                        int styleId = styleIdSupplier(cell);
                        //cell.toElement() FIXME implement
                    }
                }
            }
            zos.write(ROW_END);
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    public void end() throws IOException {
        //TODO: add here all the metadata files
        if (hasEnded) {
            throw new IllegalStateException("already ended");
        } else {
            hasEnded = true;
            writeMetadataDocuments(zos, sheets, styles, styleToIdMapping);
        }
    }


/*
    static void test() throws IOException {
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream();
             StreamingWorkbook workbook = new StreamingWorkbook(baos)) {

            workbook.withSheet("test", Stream.empty());
            workbook.withSheet("test2", Stream.empty());
            workbook.end();//<- optional if close is called
        }
    }
*/
}
