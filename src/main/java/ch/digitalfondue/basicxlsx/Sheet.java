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
import java.time.LocalDateTime;
import java.util.*;

public class Sheet {

    final SortedMap<Integer, SortedMap<Integer, Cell>> cells = new TreeMap<>();
    Map<Cell, Style> styleRegistry;

    Sheet(Map<Cell, Style> styleRegistry) {
        this.styleRegistry = styleRegistry;
    }

    public int getMaxCol() {
        int max = 0;
        for (SortedMap<Integer, Cell> e : cells.values()) {
            max = Math.max(max, e.lastKey());
        }
        return max;
    }

    public Cell setCellAt(Cell cell, int row, int column) {
        cell.styleRegistrator = styleRegistry::put;
        cell.styleRegistry = styleRegistry::get;
        cells.computeIfAbsent(row, r -> new TreeMap<>()).put(column, cell);
        return cell;
    }

    //set string
    public Cell setValueAt(String value, int row, int column) {
        return setCellAt(new Cell.StringCell(value), row, column);
    }

    //numbers
    public Cell setValueAt(long value, int row, int column) {
        return setCellAt(new Cell.NumberCell(value), row, column);
    }

    public Cell setValueAt(double value, int row, int column) {
        return setCellAt(new Cell.NumberCell(value), row, column);
    }

    public Cell setValueAt(BigDecimal value, int row, int column) {
        return setCellAt(new Cell.NumberCell(value), row, column);
    }
    //

    //boolean
    public Cell setValueAt(boolean value, int row, int column) {
        return setCellAt(new Cell.BooleanCell(value), row, column);
    }
    //


    //date
    public Cell setValueAt(Date date, int row, int column) {
        return setCellAt(new Cell.DateCell(date), row, column);
    }

    public Cell setValueAt(LocalDateTime localDateTime, int row, int column) {
        return setCellAt(new Cell.LocalDateTimeCell(localDateTime), row, column);
    }
    //

    public void removeCellAt(int row, int column) {
        if (cells.containsKey(row)) {
            Cell cell = cells.get(row).remove(column);
            if (cell != null) {
                cell.styleRegistrator = null;
                cell.styleRegistry = null;
                styleRegistry.remove(cell);
            }
        }
    }

    public Optional<Cell> getCellAt(int row, int column) {
        if (cells.containsKey(row)) {
            return Optional.ofNullable(cells.get(row).get(column));
        } else {
            return Optional.empty();
        }
    }
}
