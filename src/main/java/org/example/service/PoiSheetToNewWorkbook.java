package org.example.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.HashMap;
import java.util.Map;

public final class PoiSheetToNewWorkbook {

    private PoiSheetToNewWorkbook() {
    }

    public static void copySheetToNewWorkbook(
            Sheet sourceSheet,
            Workbook targetWorkbook,
            String targetSheetName
    ) {
        Sheet targetSheet = targetWorkbook.createSheet(targetSheetName);

        Map<CellStyle, CellStyle> styleMap = new HashMap<>();

        copyColumnWidths(sourceSheet, targetSheet);
        copyRows(sourceSheet, targetSheet, targetWorkbook, styleMap);
        copyMergedRegions(sourceSheet, targetSheet);
        copySheetDisplaySettings(sourceSheet, targetSheet);
        copyPrintSettings(sourceSheet, targetSheet, targetWorkbook);

        targetSheet.setDefaultColumnWidth(sourceSheet.getDefaultColumnWidth());
        targetSheet.setDefaultRowHeight(sourceSheet.getDefaultRowHeight());
        targetSheet.setDefaultRowHeightInPoints(sourceSheet.getDefaultRowHeightInPoints());
    }

    private static void copyRows(
            Sheet sourceSheet,
            Sheet targetSheet,
            Workbook targetWorkbook,
            Map<CellStyle, CellStyle> styleMap
    ) {
        for (int rowIndex = sourceSheet.getFirstRowNum();
             rowIndex <= sourceSheet.getLastRowNum();
             rowIndex++) {
            Row sourceRow = sourceSheet.getRow(rowIndex);
            if (sourceRow == null) {
                continue;
            }

            Row targetRow = targetSheet.createRow(rowIndex);
            targetRow.setHeight(sourceRow.getHeight());
            targetRow.setZeroHeight(sourceRow.getZeroHeight());

            copyCells(sourceRow, targetRow, targetWorkbook, styleMap);
        }
    }

    private static void copyCells(
            Row sourceRow,
            Row targetRow,
            Workbook targetWorkbook,
            Map<CellStyle, CellStyle> styleMap
    ) {
        short firstCellNum = sourceRow.getFirstCellNum();
        short lastCellNum = sourceRow.getLastCellNum();

        if (firstCellNum < 0 || lastCellNum < 0) {
            return;
        }

        for (int cellIndex = firstCellNum; cellIndex < lastCellNum; cellIndex++) {
            Cell sourceCell = sourceRow.getCell(cellIndex);
            if (sourceCell == null) {
                continue;
            }

            Cell targetCell = targetRow.createCell(cellIndex, sourceCell.getCellType());
            copyCellValue(sourceCell, targetCell);
            copyCellStyle(sourceCell, targetCell, targetWorkbook, styleMap);
        }
    }

    private static void copyCellValue(Cell sourceCell, Cell targetCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getRichStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(sourceCell)) {
                    targetCell.setCellValue(sourceCell.getLocalDateTimeCellValue());
                } else {
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case ERROR:
                targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
                break;
            case BLANK:
                targetCell.setBlank();
                break;
            default:
                break;
        }
    }

    private static void copyCellStyle(
            Cell sourceCell,
            Cell targetCell,
            Workbook targetWorkbook,
            Map<CellStyle, CellStyle> styleMap
    ) {
        CellStyle sourceStyle = sourceCell.getCellStyle();
        if (sourceStyle == null) {
            return;
        }

        CellStyle targetStyle = styleMap.get(sourceStyle);
        if (targetStyle == null) {
            targetStyle = targetWorkbook.createCellStyle();
            targetStyle.cloneStyleFrom(sourceStyle);
            styleMap.put(sourceStyle, targetStyle);
        }

        targetCell.setCellStyle(targetStyle);
    }

    private static void copyColumnWidths(Sheet sourceSheet, Sheet targetSheet) {
        int maxColumn = getMaxColumn(sourceSheet);
        for (int columnIndex = 0; columnIndex < maxColumn; columnIndex++) {
            targetSheet.setColumnWidth(columnIndex, sourceSheet.getColumnWidth(columnIndex));
            targetSheet.setColumnHidden(columnIndex, sourceSheet.isColumnHidden(columnIndex));
        }
    }

    private static void copyMergedRegions(Sheet sourceSheet, Sheet targetSheet) {
        for (int mergedRegionIndex = 0; mergedRegionIndex < sourceSheet.getNumMergedRegions(); mergedRegionIndex++) {
            CellRangeAddress mergedRegion = sourceSheet.getMergedRegion(mergedRegionIndex);
            targetSheet.addMergedRegion(mergedRegion.copy());
        }
    }

    private static void copySheetDisplaySettings(Sheet sourceSheet, Sheet targetSheet) {
        targetSheet.setDisplayFormulas(sourceSheet.isDisplayFormulas());
        targetSheet.setDisplayGridlines(sourceSheet.isDisplayGridlines());
        targetSheet.setDisplayRowColHeadings(sourceSheet.isDisplayRowColHeadings());
        targetSheet.setDisplayGuts(sourceSheet.getDisplayGuts());
        targetSheet.setDisplayZeros(sourceSheet.isDisplayZeros());
        targetSheet.setRightToLeft(sourceSheet.isRightToLeft());
    }

    private static void copyPrintSettings(Sheet sourceSheet, Sheet targetSheet, Workbook targetWorkbook) {
        targetSheet.setPrintGridlines(sourceSheet.isPrintGridlines());
        targetSheet.setPrintRowAndColumnHeadings(sourceSheet.isPrintRowAndColumnHeadings());
        targetSheet.setHorizontallyCenter(sourceSheet.getHorizontallyCenter());
        targetSheet.setVerticallyCenter(sourceSheet.getVerticallyCenter());
        targetSheet.setFitToPage(sourceSheet.getFitToPage());
        targetSheet.setAutobreaks(sourceSheet.getAutobreaks());

        for (PageMargin pageMargin : PageMargin.values()) {
            targetSheet.setMargin(pageMargin, sourceSheet.getMargin(pageMargin));
        }

        copyHeader(sourceSheet, targetSheet);
        copyFooter(sourceSheet, targetSheet);
        copyPageBreaks(sourceSheet, targetSheet);
        copyPrintSetup(sourceSheet.getPrintSetup(), targetSheet.getPrintSetup());
        copyPrintArea(sourceSheet, targetSheet, targetWorkbook);
    }

    private static void copyHeader(Sheet sourceSheet, Sheet targetSheet) {
        Header sourceHeader = sourceSheet.getHeader();
        Header targetHeader = targetSheet.getHeader();
        targetHeader.setLeft(sourceHeader.getLeft());
        targetHeader.setCenter(sourceHeader.getCenter());
        targetHeader.setRight(sourceHeader.getRight());
    }

    private static void copyFooter(Sheet sourceSheet, Sheet targetSheet) {
        Footer sourceFooter = sourceSheet.getFooter();
        Footer targetFooter = targetSheet.getFooter();
        targetFooter.setLeft(sourceFooter.getLeft());
        targetFooter.setCenter(sourceFooter.getCenter());
        targetFooter.setRight(sourceFooter.getRight());
    }

    private static void copyPageBreaks(Sheet sourceSheet, Sheet targetSheet) {
        for (int columnBreak : sourceSheet.getColumnBreaks()) {
            targetSheet.setColumnBreak(columnBreak);
        }
        for (int rowBreak : sourceSheet.getRowBreaks()) {
            targetSheet.setRowBreak(rowBreak);
        }
    }

    private static void copyPrintSetup(PrintSetup sourcePrintSetup, PrintSetup targetPrintSetup) {
        targetPrintSetup.setPaperSize(sourcePrintSetup.getPaperSize());
        targetPrintSetup.setScale(sourcePrintSetup.getScale());
        targetPrintSetup.setPageStart(sourcePrintSetup.getPageStart());
        targetPrintSetup.setFitWidth(sourcePrintSetup.getFitWidth());
        targetPrintSetup.setFitHeight(sourcePrintSetup.getFitHeight());
        targetPrintSetup.setLeftToRight(sourcePrintSetup.getLeftToRight());
        targetPrintSetup.setLandscape(sourcePrintSetup.getLandscape());
        targetPrintSetup.setValidSettings(sourcePrintSetup.getValidSettings());
        targetPrintSetup.setNoColor(sourcePrintSetup.getNoColor());
        targetPrintSetup.setDraft(sourcePrintSetup.getDraft());
        targetPrintSetup.setNotes(sourcePrintSetup.getNotes());
        targetPrintSetup.setNoOrientation(sourcePrintSetup.getNoOrientation());
        targetPrintSetup.setUsePage(sourcePrintSetup.getUsePage());
        targetPrintSetup.setHResolution(sourcePrintSetup.getHResolution());
        targetPrintSetup.setVResolution(sourcePrintSetup.getVResolution());
        targetPrintSetup.setHeaderMargin(sourcePrintSetup.getHeaderMargin());
        targetPrintSetup.setFooterMargin(sourcePrintSetup.getFooterMargin());
        targetPrintSetup.setCopies(sourcePrintSetup.getCopies());
    }

    private static void copyPrintArea(Sheet sourceSheet, Sheet targetSheet, Workbook targetWorkbook) {
        Workbook sourceWorkbook = sourceSheet.getWorkbook();
        int sourceSheetIndex = sourceWorkbook.getSheetIndex(sourceSheet);
        String printArea = sourceWorkbook.getPrintArea(sourceSheetIndex);
        if (printArea == null || printArea.isBlank()) {
            return;
        }

        targetWorkbook.setPrintArea(targetWorkbook.getSheetIndex(targetSheet), printArea);
    }

    private static int getMaxColumn(Sheet sheet) {
        int maxColumn = 0;

        for (int rowIndex = sheet.getFirstRowNum();
             rowIndex <= sheet.getLastRowNum();
             rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null && row.getLastCellNum() > maxColumn) {
                maxColumn = row.getLastCellNum();
            }
        }

        return maxColumn;
    }
}
