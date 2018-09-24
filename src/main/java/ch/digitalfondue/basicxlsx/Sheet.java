package ch.digitalfondue.basicxlsx;

import java.math.BigDecimal;
import java.util.Optional;
import java.util.SortedMap;
import java.util.TreeMap;

public class Sheet {

    final SortedMap<Integer, SortedMap<Integer, Cell>> cells = new TreeMap<>();

    public int getMaxCol() {
        int max = 0;
        for (SortedMap<Integer, Cell> e : cells.values()) {
            max = Math.max(max, e.lastKey());
        }
        return max;
    }

    public Cell setCellAt(Cell cell, int row, int column) {
        return cells.computeIfAbsent(row, r -> new TreeMap<>()).put(column, cell);
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
            cells.get(row).remove(column);
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
