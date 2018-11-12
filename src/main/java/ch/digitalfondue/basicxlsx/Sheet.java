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

    /**
     * Define an height for a given row
     *
     * @param rowIndex
     * @param height
     */
    public void setRowHeight(int rowIndex, double height) {
        this.rowHeight.put(rowIndex, height);
    }

    /**
     * Define a width for a given column
     *
     * @param columnIndex
     * @param width
     */
    public void setColumnWidth(int columnIndex, double width) {
        this.columnWidth.put(columnIndex, width);
    }

    /**
     * Define the default reading order (left to right or right to left)
     *
     * @param readingOrder
     */
    public void setReadingOrder(Style.ReadingOrder readingOrder) {
        this.readingOrder = readingOrder;
    }

    List<Cell> getColumnCells(int column) {
        List<Cell> res = new ArrayList<>();
        for (Map.Entry<Integer, SortedMap<Integer, Cell>> e : cells.entrySet()) {
            if (e.getValue() != null && e.getValue().containsKey(column) && e.getValue().get(column) != null) {
                res.add(e.getValue().get(column));
            }
        }
        return res;
    }

    /**
     * Auto resize a given colum in function of the content.
     * Note/Limitations:
     *  <ul>
     *      <li>the column will have a default minimum of 8.43 width</li>
     *      <li>at the moment, only the string type column will be considered</li>
     *      <li>the operation can consume some time, so only call it just before writing the workbook.</li>
     *  </ul>
     *
     * @param column
     */
    public void autoResizeColumn(int column) {
        Optional<Double> maxValue = getColumnCells(column)
                .stream()
                .map(CellWidthCalculator::cellWidth)
                .reduce(Math::max)
                .filter(value -> value > 8.43); //8.43 is the default length

        maxValue.ifPresent(width -> {
            setColumnWidth(column, width);
        });
    }

    /**
     * Auto resize all the column. See {@link #autoResizeColumn(int)} about the limitations.
     *
     */
    public void autoResizeAllColumns() {
        int maxColIdx = getMaxCol();
        for (int i = 0; i <= maxColIdx; i++) {
            autoResizeColumn(i);
        }
    }

    private Cell setCellAt(Cell cell, int row, int column) {
        cells.computeIfAbsent(row, r -> new TreeMap<>()).put(column, cell);
        return cell;
    }

    /**
     * Set a string cell at a given row/column.
     *
     * @param value
     * @param row
     * @param column
     * @return
     */
    public Cell setValueAt(String value, int row, int column) {
        return setCellAt(Cell.cell(value), row, column);
    }


    /**
     * Set a number at a given row/column
     *
     * @param value
     * @param row
     * @param column
     * @return
     */
    public Cell setValueAt(long value, int row, int column) {
        return setCellAt(Cell.cell(value), row, column);
    }

    /**
     * Set a number at a given row/column
     *
     * @param value
     * @param row
     * @param column
     * @return
     */
    public Cell setValueAt(double value, int row, int column) {
        return setCellAt(Cell.cell(value), row, column);
    }

    /**
     * Set a number at a given row/column
     *
     * @param value
     * @param row
     * @param column
     * @return
     */
    public Cell setValueAt(BigDecimal value, int row, int column) {
        return setCellAt(Cell.cell(value), row, column);
    }
    //


    /**
     * Set a boolean at a given row/column
     *
     * @param value
     * @param row
     * @param column
     * @return
     */
    public Cell setValueAt(boolean value, int row, int column) {
        return setCellAt(Cell.cell(value), row, column);
    }


    /**
     * Set a formula at a given row/column.
     *
     * Note: the formula name must be in english. If you target a xlsx viewer, you may need to use
     * {@link #setFormulaAt(String, String, int, int)} where you can specify the result of the formula.
     *
     * @param formula
     * @param row
     * @param column
     * @return
     */
    public Cell setFormulaAt(String formula, int row, int column) {
        return setCellAt(Cell.formula(formula), row, column);
    }


    /**
     * Set a formula at a given row/column. You can specify the result if you target a xslx viewer.
     *
     * @param formula
     * @param result
     * @param row
     * @param column
     * @return
     */
    public Cell setFormulaAt(String formula, String result, int row, int column) {
        return setCellAt(Cell.formula(formula, result), row, column);
    }
    //


    /**
     * Set a date at a given row/column. Note: you will need to specify a formatting for the date.
     *
     * @param date
     * @param row
     * @param column
     * @return
     */
    public Cell setValueAt(Date date, int row, int column) {
        return setCellAt(Cell.cell(date), row, column);
    }

    /**
     * Set a date at a given row/column. Note: you will need to specify a formatting for the date.
     *
     * @param date
     * @param row
     * @param column
     * @return
     */
    public Cell setValueAt(LocalDateTime date, int row, int column) {
        return setCellAt(Cell.cell(date), row, column);
    }

    /**
     * Set a date at a given row/column. Note: you will need to specify a formatting for the date.
     *
     * @param date
     * @param row
     * @param column
     * @return
     */
    public Cell setValueAt(LocalDate date, int row, int column) {
        return setCellAt(Cell.cell(date), row, column);
    }
    //

    /**
     * Remove a cell at a given row/column.
     *
     * @param row
     * @param column
     * @return
     */
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

    /**
     * Get a cell at a given row/column.
     *
     * @param row
     * @param column
     * @return
     */
    public Optional<Cell> getCellAt(int row, int column) {
        if (cells.containsKey(row)) {
            return Optional.ofNullable(cells.get(row).get(column));
        } else {
            return Optional.empty();
        }
    }
}
