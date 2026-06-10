package org.example.model;

public class SheetSummary {

    private final String sheetName;
    private final int physicalRowCount;
    private final int firstRowNumber;
    private final int lastRowNumber;

    public SheetSummary(String sheetName, int physicalRowCount, int firstRowNumber, int lastRowNumber) {
        this.sheetName = sheetName;
        this.physicalRowCount = physicalRowCount;
        this.firstRowNumber = firstRowNumber;
        this.lastRowNumber = lastRowNumber;
    }

    public String getSheetName() {
        return sheetName;
    }

    public int getPhysicalRowCount() {
        return physicalRowCount;
    }

    public int getFirstRowNumber() {
        return firstRowNumber;
    }

    public int getLastRowNumber() {
        return lastRowNumber;
    }
}
