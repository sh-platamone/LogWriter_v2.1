package org.example.service;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.example.model.ScheduleGenerationRequest;
import org.example.model.ScheduleGenerationResult;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class ScheduleWriterService {

    private static final List<String> DEFAULT_TIME_BLOCKS = List.of(
            "06:00 - 08:00",
            "08:00 - 10:00",
            "10:00 - 12:00",
            "12:00 - 14:00",
            "14:00 - 16:00",
            "16:00 - 18:00",
            "18:00 - 20:00",
            "20:00 - Close"
    );

    private static final DateTimeFormatter SHEET_DATE_FORMAT = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    private static final DateTimeFormatter DISPLAY_DATE_FORMAT = DateTimeFormatter.ofPattern("EEEE, MMM d yyyy");

    public ScheduleGenerationResult generate(ScheduleGenerationRequest request) throws IOException {
        Path parent = request.getOutputWorkbook().getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        try (InputStream inputStream = Files.newInputStream(request.getSourceWorkbook());
             Workbook workbook = createWorkbook(inputStream);
             OutputStream outputStream = Files.newOutputStream(request.getOutputWorkbook())) {

            String sheetName = buildGeneratedSheetName(workbook, request.getScheduleDate());
            Sheet scheduleSheet = workbook.createSheet(sheetName);
            writeScheduleSheet(workbook, scheduleSheet, request);
            workbook.setActiveSheet(workbook.getSheetIndex(scheduleSheet));
            workbook.write(outputStream);
            return new ScheduleGenerationResult(request.getOutputWorkbook(), sheetName, request.getTaskNames().size());
        }
    }

    private Workbook createWorkbook(InputStream inputStream) throws IOException {
        return WorkbookFactory.create(inputStream);
    }

    private String buildGeneratedSheetName(Workbook workbook, LocalDate scheduleDate) {
        String baseName = "Daily Schedule " + SHEET_DATE_FORMAT.format(scheduleDate);
        String sheetName = baseName;
        int suffix = 2;
        while (workbook.getSheet(sheetName) != null) {
            sheetName = baseName + " (" + suffix + ")";
            suffix++;
        }
        return sheetName;
    }

    private void writeScheduleSheet(Workbook workbook, Sheet sheet, ScheduleGenerationRequest request) {
        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle labelStyle = createLabelStyle(workbook);
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle valueStyle = createValueStyle(workbook);
        CellStyle assigneeStyle = createAssigneeStyle(workbook);

        createMergedCell(sheet, 0, 0, 0, 3, "Daily Crew Schedule", titleStyle);
        createLabelValueRow(sheet, 2, "Store / Team", fallback(request.getStoreName(), "Grocery Crew"), labelStyle, valueStyle);
        createLabelValueRow(sheet, 3, "Schedule Date", DISPLAY_DATE_FORMAT.format(request.getScheduleDate()), labelStyle, valueStyle);
        createLabelValueRow(sheet, 4, "Source Sheet", request.getSourceSheetName(), labelStyle, valueStyle);
        createLabelValueRow(sheet, 5, "Notes", fallback(request.getNotes(), "No notes supplied."), labelStyle, valueStyle);

        Row headerRow = sheet.createRow(7);
        createCell(headerRow, 0, "Time Block", headerStyle);
        createCell(headerRow, 1, "Task", headerStyle);
        createCell(headerRow, 2, "Assigned Crew", headerStyle);
        createCell(headerRow, 3, "Notes", headerStyle);

        List<String> tasks = request.getTaskNames().isEmpty()
                ? List.of("Unassigned task")
                : request.getTaskNames();

        for (int index = 0; index < tasks.size(); index++) {
            Row row = sheet.createRow(8 + index);
            createCell(row, 0, resolveTimeBlock(index), valueStyle);
            createCell(row, 1, tasks.get(index), valueStyle);
            createCell(row, 2, "", assigneeStyle);
            createCell(row, 3, "", assigneeStyle);
        }

        sheet.createFreezePane(0, 8);
        sheet.setAutoFilter(new CellRangeAddress(7, 7, 0, 3));
        for (int columnIndex = 0; columnIndex <= 3; columnIndex++) {
            sheet.autoSizeColumn(columnIndex);
            sheet.setColumnWidth(columnIndex, Math.max(sheet.getColumnWidth(columnIndex), 18 * 256));
        }
    }

    private void createLabelValueRow(Sheet sheet, int rowIndex, String label, String value, CellStyle labelStyle, CellStyle valueStyle) {
        Row row = sheet.createRow(rowIndex);
        createCell(row, 0, label, labelStyle);
        createCell(row, 1, value, valueStyle);
        createCell(row, 2, "", valueStyle);
        createCell(row, 3, "", valueStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 3));
    }

    private void createMergedCell(Sheet sheet, int rowIndex, int columnIndex, int mergeToRow, int mergeToColumn, String value, CellStyle style) {
        Row row = sheet.createRow(rowIndex);
        createCell(row, columnIndex, value, style);
        for (int nextColumn = columnIndex + 1; nextColumn <= mergeToColumn; nextColumn++) {
            createCell(row, nextColumn, "", style);
        }
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, mergeToRow, columnIndex, mergeToColumn));
    }

    private Cell createCell(Row row, int columnIndex, String value, CellStyle style) {
        Cell cell = row.createCell(columnIndex);
        cell.setCellValue(value);
        cell.setCellStyle(style);
        return cell;
    }

    private String resolveTimeBlock(int index) {
        if (index < DEFAULT_TIME_BLOCKS.size()) {
            return DEFAULT_TIME_BLOCKS.get(index);
        }
        return "Flexible Block " + (index - DEFAULT_TIME_BLOCKS.size() + 1);
    }

    private String fallback(String value, String fallback) {
        return value == null || value.isBlank() ? fallback : value;
    }

    private CellStyle createTitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.DARK_GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private CellStyle createLabelStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private CellStyle createValueStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private CellStyle createAssigneeStyle(Workbook workbook) {
        CellStyle style = createValueStyle(workbook);
        style.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }
}
