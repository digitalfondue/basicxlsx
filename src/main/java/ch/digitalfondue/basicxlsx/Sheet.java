package ch.digitalfondue.basicxlsx;

import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

public class Sheet {

    private final String name;

    private final Map<Integer, Map<Integer, Cell>> cells = new HashMap<>();

    Sheet(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }

    public void setCellAt(Cell cell, int row, int column) {
        cells.computeIfAbsent(row, r -> new HashMap<>()).put(column, cell);
    }

    public void removeCellAt(int row, int column) {
        if(cells.containsKey(row)) {
            cells.get(row).remove(column);
        }
    }

    public Optional<Cell> getCellAt(int row, int column) {
        if(cells.containsKey(row)) {
            return Optional.ofNullable(cells.get(row).get(column));
        } else {
            return Optional.empty();
        }
    }
}
