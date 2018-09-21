package ch.digitalfondue.basicxlsx;

import java.util.HashMap;
import java.util.Map;

public class Sheet {

    private final String name;

    private final Map<Integer, Map<Integer, Cell>> cells = new HashMap<>();

    Sheet(String name) {
        this.name = name;
    }

    public void setCellAt(Cell cell, int row, int column) {
        cells.computeIfAbsent(row, r -> new HashMap<>()).put(column, cell);
    }
}
