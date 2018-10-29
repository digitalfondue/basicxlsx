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

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

/**
 * Represent an xlsx sheet.
 */
public class Sheet {

    final SortedMap<Integer, SortedMap<Integer, Cell>> cells = new TreeMap<>();
    final Map<Integer, Double> rowHeight = new HashMap<>();
    final Map<Integer, Double> columnWidth = new HashMap<>();
    Style.ReadingOrder readingOrder;

    Sheet() {
    }

    int getMaxCol() {
        int max = 0;
        for (SortedMap<Integer, Cell> e : cells.values()) {
            max = Math.max(max, e.lastKey());
        }
        return max;
    }

    public void setRowHeight(int rowIndex, double height) {
        this.rowHeight.put(rowIndex, height);
    }

    public void setColumnWidth(int columnIndex, double width) {
        this.columnWidth.put(columnIndex, width);
    }

    public void setReadingOrder(Style.ReadingOrder readingOrder) {
        this.readingOrder = readingOrder;
    }

    private Cell setCellAt(Cell cell, int row, int column) {
        cells.computeIfAbsent(row, r -> new TreeMap<>()).put(column, cell);
        return cell;
    }

    //set string
    public Cell setValueAt(String value, int row, int column) {
        return setCellAt(Cell.cell(value), row, column);
    }

    //numbers
    public Cell setValueAt(long value, int row, int column) {
        return setCellAt(Cell.cell(value), row, column);
    }

    public Cell setValueAt(double value, int row, int column) {
        return setCellAt(Cell.cell(value), row, column);
    }

    public Cell setValueAt(BigDecimal value, int row, int column) {
        return setCellAt(Cell.cell(value), row, column);
    }
    //

    //boolean
    public Cell setValueAt(boolean value, int row, int column) {
        return setCellAt(Cell.cell(value), row, column);
    }
    //

    //formula
    public Cell setFormulaAt(String formula, int row, int column) {
        return setCellAt(Cell.formula(formula), row, column);
    }

    public Cell setFormulaAt(String formula, String result, int row, int column) {
        return setCellAt(Cell.formula(formula, result), row, column);
    }
    //


    //date
    public Cell setValueAt(Date date, int row, int column) {
        return setCellAt(Cell.cell(date), row, column);
    }

    public Cell setValueAt(LocalDateTime localDateTime, int row, int column) {
        return setCellAt(Cell.cell(localDateTime), row, column);
    }

    public Cell setValueAt(LocalDate localDate, int row, int column) {
        return setCellAt(Cell.cell(localDate), row, column);
    }
    //

    public Optional<Cell> removeCellAt(int row, int column) {
        Cell cell = null;
        if (cells.containsKey(row)) {
            cell = cells.get(row).remove(column);
            if (cell != null) {
                cell.style = null;
            }
        }
        return Optional.ofNullable(cell);
    }

    public Optional<Cell> getCellAt(int row, int column) {
        if (cells.containsKey(row)) {
            return Optional.ofNullable(cells.get(row).get(column));
        } else {
            return Optional.empty();
        }
    }
}
