package org.example.model;

public class DaySheetOption {

    private final String dayLabel;
    private final int sheetIndex;
    private final String sheetName;

    public DaySheetOption(String dayLabel, int sheetIndex, String sheetName) {
        this.dayLabel = dayLabel;
        this.sheetIndex = sheetIndex;
        this.sheetName = sheetName;
    }

    public String getDayLabel() {
        return dayLabel;
    }

    public int getSheetIndex() {
        return sheetIndex;
    }

    public String getSheetName() {
        return sheetName;
    }

    @Override
    public String toString() {
        return dayLabel;
    }
}
