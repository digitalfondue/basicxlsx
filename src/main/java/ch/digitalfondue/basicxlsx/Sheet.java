package ch.digitalfondue.basicxlsx;

import java.util.Optional;
import java.util.SortedMap;
import java.util.TreeMap;

public class Sheet {

    private final String name;

    final SortedMap<Integer, SortedMap<Integer, Cell>> cells = new TreeMap<>();

    Sheet(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }

    public int getMaxCol() {
        int max = 0;
        for (SortedMap<Integer, Cell> e : cells.values()) {
            max = Math.max(max, e.lastKey());
        }
        return max;
    }

    public void setCellAt(Cell cell, int row, int column) {
        cells.computeIfAbsent(row, r -> new TreeMap<>()).put(column, cell);
    }

    //set string
    public void setValueAt(String value, int row, int column) {
        setCellAt(new Cell.StringCell(value), row, column);
    }

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
