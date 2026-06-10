package org.example.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.example.model.RegisterBlockKey;
import org.example.model.RegisterSheetState;
import org.example.model.CrewAvailabilitySummary;

import java.util.ArrayList;
import java.util.List;

public class CrewAvailabilityStatusService {

    private static final String LEADERSHIP_TEAM_MARKER = "Leadership Team";

    public CrewAvailabilitySummary summarizeAvailability(
            Workbook workbook,
            Sheet sheet,
            RegisterSheetState registerSheetState,
            int requestedStartHour,
            int requestedEndHour
    ) {
        if (workbook == null || sheet == null) {
            return new CrewAvailabilitySummary(0, 0, 0, null, 0, null);
        }

        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        int leadershipTeamRowIndex = findLeadershipTeamRowIndex(sheet, formatter, evaluator);
        List<Integer> crewRows = collectCrewRows(sheet, leadershipTeamRowIndex, formatter, evaluator);
        if (crewRows.isEmpty()) {
            return new CrewAvailabilitySummary(0, 0, 0, null, 0, null);
        }

        int openerCount = countAvailableCrewForHour(
                sheet,
                crewRows,
                RegisterGridLayout.FIRST_REGISTER_HOUR,
                formatter,
                evaluator,
                registerSheetState,
                true
        );
        int closerCount = countAvailableCrewForHour(
                sheet,
                crewRows,
                RegisterGridLayout.LAST_REGISTER_HOUR,
                formatter,
                evaluator,
                registerSheetState,
                true
        );

        int startHour = RegisterGridLayout.clampHour(Math.min(requestedStartHour, requestedEndHour));
        int endHour = RegisterGridLayout.clampHour(Math.max(requestedStartHour, requestedEndHour));
        int maximumAvailableCrewCount = -1;
        Integer maximumAvailableHour = null;
        int minimumAvailableCrewCount = Integer.MAX_VALUE;
        Integer minimumAvailableHour = null;

        for (int hour = startHour; hour <= endHour; hour++) {
            int availableCrewCount = countAvailableCrewForHour(
                    sheet,
                    crewRows,
                    hour,
                    formatter,
                    evaluator,
                    registerSheetState,
                    false
            );
            if (availableCrewCount > maximumAvailableCrewCount) {
                maximumAvailableCrewCount = availableCrewCount;
                maximumAvailableHour = hour;
            }
            if (availableCrewCount < minimumAvailableCrewCount) {
                minimumAvailableCrewCount = availableCrewCount;
                minimumAvailableHour = hour;
            }
        }

        if (maximumAvailableHour == null || minimumAvailableHour == null) {
            return new CrewAvailabilitySummary(openerCount, closerCount, 0, null, 0, null);
        }

        return new CrewAvailabilitySummary(
                openerCount,
                closerCount,
                maximumAvailableCrewCount,
                maximumAvailableHour,
                minimumAvailableCrewCount,
                minimumAvailableHour
        );
    }

    private int countAvailableCrewForHour(
            Sheet sheet,
            List<Integer> crewRows,
            int hour,
            DataFormatter formatter,
            FormulaEvaluator evaluator,
            RegisterSheetState registerSheetState,
            boolean ignoreRegisterAssignments
    ) {
        int firstColumnIndex = RegisterGridLayout.hourToFirstColumn(hour);
        int availableCrewCount = 0;
        for (int rowIndex : crewRows) {
            if (isHourBlockAvailable(
                    sheet,
                    rowIndex,
                    firstColumnIndex,
                    formatter,
                    evaluator,
                    registerSheetState,
                    ignoreRegisterAssignments
            )) {
                availableCrewCount++;
            }
        }
        return availableCrewCount;
    }

    private boolean isHourBlockAvailable(
            Sheet sheet,
            int rowIndex,
            int firstColumnIndex,
            DataFormatter formatter,
            FormulaEvaluator evaluator,
            RegisterSheetState registerSheetState,
            boolean ignoreRegisterAssignments
    ) {
        RegisterBlockKey registerBlockKey = new RegisterBlockKey(rowIndex, firstColumnIndex);
        return isCellAvailable(
                sheet,
                rowIndex,
                firstColumnIndex,
                formatter,
                evaluator,
                registerSheetState,
                ignoreRegisterAssignments,
                registerBlockKey
        ) && isCellAvailable(
                sheet,
                rowIndex,
                firstColumnIndex + 1,
                formatter,
                evaluator,
                registerSheetState,
                ignoreRegisterAssignments,
                registerBlockKey
        );
    }

    private boolean isCellAvailable(
            Sheet sheet,
            int rowIndex,
            int columnIndex,
            DataFormatter formatter,
            FormulaEvaluator evaluator,
            RegisterSheetState registerSheetState,
            boolean ignoreRegisterAssignments,
            RegisterBlockKey registerBlockKey
    ) {
        Cell cell = getCell(sheet, rowIndex, columnIndex);
        if (cell == null) {
            return hasNoFillStyle(null, rowIndex, columnIndex, registerSheetState, ignoreRegisterAssignments, registerBlockKey);
        }

        String displayValue = formatter.formatCellValue(cell, evaluator);
        if (!displayValue.isBlank()) {
            return false;
        }

        CellStyle cellStyle = cell.getCellStyle();
        return hasNoFillStyle(
                cellStyle,
                rowIndex,
                columnIndex,
                registerSheetState,
                ignoreRegisterAssignments,
                registerBlockKey
        );
    }

    private boolean hasNoFillStyle(
            CellStyle liveCellStyle,
            int rowIndex,
            int columnIndex,
            RegisterSheetState registerSheetState,
            boolean ignoreRegisterAssignments,
            RegisterBlockKey registerBlockKey
    ) {
        CellStyle effectiveStyle = liveCellStyle;
        if (ignoreRegisterAssignments
                && registerSheetState != null
                && registerSheetState.isAssigned(registerBlockKey)) {
            effectiveStyle = registerSheetState.getBaselineStyle(rowIndex, columnIndex);
        }

        return effectiveStyle == null || effectiveStyle.getFillPattern() == FillPatternType.NO_FILL;
    }

    private List<Integer> collectCrewRows(
            Sheet sheet,
            int leadershipTeamRowIndex,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        List<Integer> crewRows = new ArrayList<>();
        int lastCrewRowIndex = Math.min(sheet.getLastRowNum(), leadershipTeamRowIndex - 1);
        for (int rowIndex = RegisterGridLayout.HOUR_GRID_START_ROW_INDEX; rowIndex <= lastCrewRowIndex; rowIndex++) {
            if (isCrewRow(sheet, rowIndex, leadershipTeamRowIndex, formatter, evaluator)) {
                crewRows.add(rowIndex);
            }
        }
        return crewRows;
    }

    private boolean isCrewRow(
            Sheet sheet,
            int rowIndex,
            int leadershipTeamRowIndex,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        if (rowIndex >= leadershipTeamRowIndex) {
            return false;
        }

        Row row = sheet.getRow(rowIndex);
        if (row == null || row.getZeroHeight()) {
            return false;
        }

        for (int columnIndex = 0; columnIndex < RegisterGridLayout.HOUR_GRID_START_COLUMN_INDEX; columnIndex++) {
            Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
            if (cell != null && !formatter.formatCellValue(cell, evaluator).isBlank()) {
                return true;
            }
        }

        return false;
    }

    private int findLeadershipTeamRowIndex(
            Sheet sheet,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        for (int rowIndex = RegisterGridLayout.HOUR_GRID_START_ROW_INDEX; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }

            short lastCellNumber = row.getLastCellNum();
            if (lastCellNumber < 0) {
                continue;
            }

            for (int columnIndex = 0; columnIndex < lastCellNumber; columnIndex++) {
                Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                if (cell == null) {
                    continue;
                }

                String displayValue = formatter.formatCellValue(cell, evaluator).trim();
                if (LEADERSHIP_TEAM_MARKER.equals(displayValue)) {
                    return rowIndex;
                }
            }
        }

        return sheet.getLastRowNum() + 1;
    }

    private Cell getCell(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return null;
        }
        return row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
    }
}
