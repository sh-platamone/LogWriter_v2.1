package org.example.model;

public class RowPreview {

    private final int rowNumber;
    private final int cellCount;
    private final String summary;

    public RowPreview(int rowNumber, int cellCount, String summary) {
        this.rowNumber = rowNumber;
        this.cellCount = cellCount;
        this.summary = summary;
    }

    public int getRowNumber() {
        return rowNumber;
    }

    public int getCellCount() {
        return cellCount;
    }

    public String getSummary() {
        return summary;
    }
}
