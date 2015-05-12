package com.aspose.spreadsheeteditor;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 *
 * @author saqib
 */
public class Row implements Serializable {

    private int id;
    private Map<String, Cell> cellsMap = new HashMap<>();
    private List<Cell> cellsList = new ArrayList<>();

    public int getId() {
        return id;
    }

    public Map<String, Cell> getCellsMap() {
        return this.cellsMap;
    }

    public void putCell(String columnName, Cell c) {
        this.cellsMap.put(columnName, c);
    }

    public void putCell(int columnId, Cell c) {
        this.cellsMap.put(com.aspose.cells.CellsHelper.columnIndexToName(columnId), c);
    }

    public List<Cell> getCellsList() {
        return this.cellsList;
    }

    public static class Builder {

        protected Row instance;

        public Builder() {
            this.instance = new Row();
        }

        public int getId() {
            return this.instance.id;
        }

        public Builder setId(int id) {
            this.instance.id = id;
            return this;
        }

        public Builder setCell(String column, Cell cell) {
            this.instance.cellsMap.put(column, cell);

            int columnId = com.aspose.cells.CellsHelper.columnNameToIndex(column);
            this.instance.cellsList.add(columnId, cell);
            return this;
        }

        public Row build() {
            return instance;
        }
    }
}
