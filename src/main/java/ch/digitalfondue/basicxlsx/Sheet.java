package ch.digitalfondue.basicxlsx;

import java.math.BigDecimal;
import java.util.Map;
import java.util.Optional;
import java.util.SortedMap;
import java.util.TreeMap;

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
