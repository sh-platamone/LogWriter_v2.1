package org.example.service;

public final class RegisterGridLayout {

    public static final int HOUR_GRID_START_ROW_INDEX = 2;
    public static final int HOUR_GRID_START_COLUMN_INDEX = 10;
    public static final int HOUR_GRID_END_COLUMN_EXCLUSIVE = 36;
    public static final int FIRST_REGISTER_HOUR = 8;
    public static final int LAST_REGISTER_HOUR = FIRST_REGISTER_HOUR
            + ((HOUR_GRID_END_COLUMN_EXCLUSIVE - HOUR_GRID_START_COLUMN_INDEX) / 2) - 1;

    private RegisterGridLayout() {
    }

    public static boolean isHourGridCoordinate(int rowIndex, int columnIndex) {
        return isHourGridRow(rowIndex) && isHourGridColumn(columnIndex);
    }

    public static boolean isHourGridRow(int rowIndex) {
        return rowIndex >= HOUR_GRID_START_ROW_INDEX;
    }

    public static boolean isHourGridColumn(int columnIndex) {
        return columnIndex >= HOUR_GRID_START_COLUMN_INDEX
                && columnIndex < HOUR_GRID_END_COLUMN_EXCLUSIVE;
    }

    public static int normalizeHourBlockStartColumn(int columnIndex) {
        if (!isHourGridColumn(columnIndex)) {
            throw new IllegalArgumentException("Column %d is outside the register hour grid.".formatted(columnIndex));
        }

        int offset = columnIndex - HOUR_GRID_START_COLUMN_INDEX;
        return HOUR_GRID_START_COLUMN_INDEX + ((offset / 2) * 2);
    }

    public static int hourToFirstColumn(int hour) {
        int clampedHour = clampHour(hour);
        return HOUR_GRID_START_COLUMN_INDEX + ((clampedHour - FIRST_REGISTER_HOUR) * 2);
    }

    public static int firstColumnToHour(int firstColumnIndex) {
        if (!isHourGridColumn(firstColumnIndex)
                || ((firstColumnIndex - HOUR_GRID_START_COLUMN_INDEX) % 2) != 0) {
            throw new IllegalArgumentException(
                    "Column %d is not the start of a register hour block.".formatted(firstColumnIndex)
            );
        }

        return FIRST_REGISTER_HOUR + ((firstColumnIndex - HOUR_GRID_START_COLUMN_INDEX) / 2);
    }

    public static int clampHour(int hour) {
        return Math.max(FIRST_REGISTER_HOUR, Math.min(LAST_REGISTER_HOUR, hour));
    }
}
