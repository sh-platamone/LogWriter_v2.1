package org.example.service;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Bounds;
import javafx.scene.Node;
import javafx.scene.control.ScrollBar;
import javafx.geometry.Orientation;
import javafx.stage.Stage;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
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
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetCellType;
import org.controlsfx.control.spreadsheet.SpreadsheetColumn;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.BitSet;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class WorkbookSpreadsheetViewService {

    private static final int TRAILING_COLUMNS_TO_REMOVE = 2;
    private static final String GAP_SCALING_ENABLED_KEY = "logwriter-gap-scaling-enabled";
    private static final int GAP_CLOSING_PASSES = 6;
    private static final String BASE_GRIDLINE_COLOR = "#D4D8DB";
    private static final String BASE_GRIDLINE_WIDTH = "0.28";
    private static final double MIN_RENDERED_COLUMN_WIDTH = 8.0;

    public SpreadsheetView createSpreadsheetView(Path workbookPath, int sheetIndex) throws IOException {
        try (InputStream inputStream = Files.newInputStream(workbookPath);
             Workbook workbook = WorkbookFactory.create(inputStream)) {
            return createSpreadsheetView(workbook, sheetIndex);
        }
    }

    public SpreadsheetView createSpreadsheetView(Workbook workbook, int sheetIndex) {
        if (sheetIndex < 0 || sheetIndex >= workbook.getNumberOfSheets()) {
            throw new IllegalArgumentException("Sheet index %d is out of bounds.".formatted(sheetIndex));
        }

        Sheet sheet = workbook.getSheetAt(sheetIndex);
        SpreadsheetRenderState renderState = buildRenderState(workbook, sheet);
        GridBase grid = new GridBase(renderState.rowCount(), renderState.columnCount());
        grid.setRows(renderState.rows());
        grid.getRowHeaders().setAll(renderState.rowHeaders());
        grid.getColumnHeaders().setAll(renderState.columnHeaders());
        grid.setRowHeightCallback(rowIndex -> renderState.rowHeights().getOrDefault(rowIndex, renderState.defaultRowHeight()));

        for (CellRangeAddress mergedRegion : renderState.mergedRegions()) {
            int rowSpan = mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
            int columnSpan = mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1;
            if (rowSpan > 1) {
                grid.spanRow(rowSpan, mergedRegion.getFirstRow(), mergedRegion.getFirstColumn());
            }
            if (columnSpan > 1) {
                grid.spanColumn(columnSpan, mergedRegion.getFirstRow(), mergedRegion.getFirstColumn());
            }
        }

        SpreadsheetView spreadsheetView = new SpreadsheetView(grid);
        spreadsheetView.setEditable(true);
        spreadsheetView.setShowRowHeader(true);
        spreadsheetView.setShowColumnHeader(true);
        spreadsheetView.setFixingRowsAllowed(true);
        spreadsheetView.setFixingColumnsAllowed(true);
        spreadsheetView.setRowHeaderWidth(52);
        spreadsheetView.setHiddenRows(renderState.hiddenRows());
        spreadsheetView.setHiddenColumns(renderState.hiddenColumns());

        applyViewState(spreadsheetView, renderState);
        enableGapClosingColumnScaling(spreadsheetView, renderState.hiddenColumns());
        return spreadsheetView;
    }

    public void refreshRegisterBlockDisplay(
            SpreadsheetView spreadsheetView,
            Workbook workbook,
            int sheetIndex,
            int rowIndex,
            int firstColumnIndex
    ) {
        if (spreadsheetView == null || workbook == null) {
            return;
        }
        if (sheetIndex < 0 || sheetIndex >= workbook.getNumberOfSheets()) {
            return;
        }

        Sheet sheet = workbook.getSheetAt(sheetIndex);
        int rowCount = Math.max(sheet.getLastRowNum() + 1, 1);
        int columnCount = determineRenderableColumnCount(sheet);
        if (rowIndex < 0 || rowIndex >= rowCount || firstColumnIndex < 0 || firstColumnIndex >= columnCount) {
            return;
        }

        List<CellRangeAddress> mergedRegions = collectRenderMergedRegions(workbook, sheet, rowCount, columnCount);
        refreshDisplayedCellStyle(spreadsheetView, workbook, sheet, mergedRegions, rowCount, columnCount, rowIndex, firstColumnIndex);

        int secondColumnIndex = firstColumnIndex + 1;
        if (secondColumnIndex < columnCount
                && spreadsheetView.getGrid().getRows().get(rowIndex).get(firstColumnIndex)
                != spreadsheetView.getGrid().getRows().get(rowIndex).get(secondColumnIndex)) {
            refreshDisplayedCellStyle(
                    spreadsheetView,
                    workbook,
                    sheet,
                    mergedRegions,
                    rowCount,
                    columnCount,
                    rowIndex,
                    secondColumnIndex
            );
        }
    }

    private SpreadsheetRenderState buildRenderState(Workbook workbook, Sheet sheet) {
        int rowCount = Math.max(sheet.getLastRowNum() + 1, 1);
        int columnCount = determineRenderableColumnCount(sheet);
        DataFormatter formatter = new DataFormatter();
        var evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        ObservableList<ObservableList<SpreadsheetCell>> rows = FXCollections.observableArrayList();
        Map<Integer, Double> rowHeights = new HashMap<>();
        List<String> rowHeaders = new ArrayList<>();
        List<String> columnHeaders = new ArrayList<>();
        List<Double> columnWidths = new ArrayList<>();
        BitSet hiddenRows = new BitSet(rowCount);
        BitSet hiddenColumns = new BitSet(columnCount);
        List<CellRangeAddress> mergedRegions = new ArrayList<>();
        Map<String, CellRangeAddress> mergedRegionAnchors;
        double defaultRowHeight = rowHeightPointsToPixels(sheet.getDefaultRowHeightInPoints());

        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            columnHeaders.add(CellReference.convertNumToColString(columnIndex));
            columnWidths.add((double) Math.max(24f, sheet.getColumnWidthInPixels(columnIndex)));
            if (sheet.isColumnHidden(columnIndex)) {
                hiddenColumns.set(columnIndex);
            }
        }

        mergedRegions.addAll(collectRenderMergedRegions(workbook, sheet, rowCount, columnCount));
        mergedRegionAnchors = indexMergedRegionAnchors(mergedRegions);

        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            ObservableList<SpreadsheetCell> spreadsheetRow = FXCollections.observableArrayList();
            rows.add(spreadsheetRow);
            rowHeaders.add(Integer.toString(rowIndex + 1));

            if (row != null) {
                rowHeights.put(rowIndex, rowHeightPointsToPixels(row.getHeightInPoints()));
                if (row.getZeroHeight()) {
                    hiddenRows.set(rowIndex);
                }
            } else {
                rowHeights.put(rowIndex, defaultRowHeight);
            }

            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                Cell cell = row == null ? null : row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                CellRangeAddress mergedRegion = mergedRegionAnchors.get(cellKey(rowIndex, columnIndex));
                int rowSpan = mergedRegion == null ? 1 : mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
                int columnSpan = mergedRegion == null ? 1 : mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1;
                String displayValue = cell == null ? "" : formatter.formatCellValue(cell, evaluator);
                SpreadsheetCell spreadsheetCell = SpreadsheetCellType.STRING.createCell(rowIndex, columnIndex, rowSpan, columnSpan, displayValue);
                spreadsheetCell.setEditable(true);

                if (cell != null) {
                    CellStyle cellStyle = cell.getCellStyle();
                    spreadsheetCell.setWrapText(cellStyle.getWrapText());
                    String style = buildCellStyle(workbook, sheet, rowIndex, columnIndex, rowCount, columnCount, mergedRegion, cellStyle);
                    if (!style.isBlank()) {
                        spreadsheetCell.setStyle(style);
                    }
                }

                spreadsheetRow.add(spreadsheetCell);
            }
        }
        applyMergedCellReferences(rows, mergedRegions, columnCount);

        PaneInformation paneInformation = sheet.getPaneInformation();
        int frozenRowCount = 0;
        int frozenColumnCount = 0;
        if (paneInformation != null && paneInformation.isFreezePane()) {
            frozenRowCount = paneInformation.getHorizontalSplitTopRow();
            frozenColumnCount = paneInformation.getVerticalSplitLeftColumn();
        }

        return new SpreadsheetRenderState(
                rows,
                rowCount,
                columnCount,
                rowHeights,
                defaultRowHeight,
                rowHeaders,
                columnHeaders,
                columnWidths,
                hiddenRows,
                hiddenColumns,
                mergedRegions,
                frozenRowCount,
                frozenColumnCount
        );
    }

    private int determineRenderableColumnCount(Sheet sheet) {
        int columnCount = 1;
        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null && row.getLastCellNum() > 0) {
                columnCount = Math.max(columnCount, row.getLastCellNum());
            }
        }

        for (int index = 0; index < sheet.getNumMergedRegions(); index++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(index);
            columnCount = Math.max(columnCount, mergedRegion.getLastColumn() + 1);
        }
        return Math.max(1, columnCount - TRAILING_COLUMNS_TO_REMOVE);
    }

    private List<CellRangeAddress> collectRenderMergedRegions(
            Workbook workbook,
            Sheet sheet,
            int rowCount,
            int columnCount
    ) {
        List<CellRangeAddress> mergedRegions = new ArrayList<>();
        for (int index = 0; index < sheet.getNumMergedRegions(); index++) {
            CellRangeAddress clippedRegion = clipMergedRegion(sheet.getMergedRegion(index), rowCount, columnCount);
            if (clippedRegion != null) {
                mergedRegions.add(clippedRegion);
            }
        }
        appendManualHourPairSpans(workbook, sheet, rowCount, columnCount, mergedRegions);
        return mergedRegions;
    }

    private CellRangeAddress clipMergedRegion(CellRangeAddress mergedRegion, int rowCount, int columnCount) {
        if (mergedRegion.getFirstRow() >= rowCount || mergedRegion.getFirstColumn() >= columnCount) {
            return null;
        }

        int lastRow = Math.min(mergedRegion.getLastRow(), rowCount - 1);
        int lastColumn = Math.min(mergedRegion.getLastColumn(), columnCount - 1);
        if (lastRow < mergedRegion.getFirstRow() || lastColumn < mergedRegion.getFirstColumn()) {
            return null;
        }

        return new CellRangeAddress(
                mergedRegion.getFirstRow(),
                lastRow,
                mergedRegion.getFirstColumn(),
                lastColumn
        );
    }

    private void appendManualHourPairSpans(
            Workbook workbook,
            Sheet sheet,
            int rowCount,
            int columnCount,
            List<CellRangeAddress> mergedRegions
    ) {
        DataFormatter formatter = new DataFormatter();
        var evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        int endingColumnExclusive = Math.min(columnCount, RegisterGridLayout.HOUR_GRID_END_COLUMN_EXCLUSIVE);
        for (int rowIndex = RegisterGridLayout.HOUR_GRID_START_ROW_INDEX; rowIndex < rowCount; rowIndex++) {
            for (int columnIndex = RegisterGridLayout.HOUR_GRID_START_COLUMN_INDEX;
                 columnIndex + 1 < endingColumnExclusive;
                 columnIndex += 2) {

                if (intersectsMergedRegion(mergedRegions, rowIndex, columnIndex, columnIndex + 1)) {
                    continue;
                }
                if (!shouldMergeHourPair(sheet, workbook, rowIndex, columnIndex, formatter, evaluator)) {
                    continue;
                }

                mergedRegions.add(new CellRangeAddress(
                        rowIndex,
                        rowIndex,
                        columnIndex,
                        columnIndex + 1
                ));
            }
        }
    }

    private boolean intersectsMergedRegion(
            List<CellRangeAddress> mergedRegions,
            int rowIndex,
            int firstColumnIndex,
            int lastColumnIndex
    ) {
        for (CellRangeAddress mergedRegion : mergedRegions) {
            if (mergedRegion.getFirstRow() <= rowIndex
                    && mergedRegion.getLastRow() >= rowIndex
                    && mergedRegion.getFirstColumn() <= lastColumnIndex
                    && mergedRegion.getLastColumn() >= firstColumnIndex) {
                return true;
            }
        }
        return false;
    }

    private boolean shouldMergeHourPair(
            Sheet sheet,
            Workbook workbook,
            int rowIndex,
            int firstColumnIndex,
            DataFormatter formatter,
            org.apache.poi.ss.usermodel.FormulaEvaluator evaluator
    ) {
        Cell firstCell = getCell(sheet, rowIndex, firstColumnIndex);
        Cell secondCell = getCell(sheet, rowIndex, firstColumnIndex + 1);
        String firstDisplayValue = firstCell == null ? "" : formatter.formatCellValue(firstCell, evaluator);
        String secondDisplayValue = secondCell == null ? "" : formatter.formatCellValue(secondCell, evaluator);
        if (!firstDisplayValue.equals(secondDisplayValue)) {
            return false;
        }

        CellStyle firstStyle = firstCell == null ? null : firstCell.getCellStyle();
        CellStyle secondStyle = secondCell == null ? null : secondCell.getCellStyle();
        return fillSignature(workbook, firstStyle).equals(fillSignature(workbook, secondStyle));
    }

    private String fillSignature(Workbook workbook, CellStyle cellStyle) {
        if (cellStyle == null || cellStyle.getFillPattern() == FillPatternType.NO_FILL) {
            return FillPatternType.NO_FILL.name();
        }

        return cellStyle.getFillPattern().name() + ':' + orDefault(resolveFillColor(cellStyle, workbook));
    }

    private Map<String, CellRangeAddress> indexMergedRegionAnchors(List<CellRangeAddress> mergedRegions) {
        Map<String, CellRangeAddress> anchorMap = new HashMap<>();
        for (CellRangeAddress mergedRegion : mergedRegions) {
            anchorMap.put(cellKey(mergedRegion.getFirstRow(), mergedRegion.getFirstColumn()), mergedRegion);
        }
        return anchorMap;
    }

    private String cellKey(int rowIndex, int columnIndex) {
        return rowIndex + ":" + columnIndex;
    }

    private void applyMergedCellReferences(
            ObservableList<ObservableList<SpreadsheetCell>> rows,
            List<CellRangeAddress> mergedRegions,
            int columnCount
    ) {
        for (CellRangeAddress mergedRegion : mergedRegions) {
            int firstRow = mergedRegion.getFirstRow();
            int lastRow = Math.min(mergedRegion.getLastRow(), rows.size() - 1);
            int firstColumn = mergedRegion.getFirstColumn();
            int lastColumn = Math.min(mergedRegion.getLastColumn(), columnCount - 1);

            if (firstRow < 0 || firstRow >= rows.size() || firstColumn < 0 || firstColumn >= columnCount) {
                continue;
            }

            SpreadsheetCell topLeftCell = rows.get(firstRow).get(firstColumn);
            for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
                ObservableList<SpreadsheetCell> row = rows.get(rowIndex);
                for (int columnIndex = firstColumn; columnIndex <= lastColumn; columnIndex++) {
                    row.set(columnIndex, topLeftCell);
                }
            }
        }
    }

    private void applyViewState(SpreadsheetView spreadsheetView, SpreadsheetRenderState renderState) {
        List<SpreadsheetColumn> columns = spreadsheetView.getColumns();
        for (int columnIndex = 0; columnIndex < Math.min(columns.size(), renderState.columnWidths().size()); columnIndex++) {
            SpreadsheetColumn column = columns.get(columnIndex);
            double baseWidth = renderState.columnWidths().get(columnIndex);
            column.setPrefWidth(baseWidth);
            column.setMinWidth(MIN_RENDERED_COLUMN_WIDTH);
            column.setMaxWidth(Double.MAX_VALUE);
            column.setResizable(true);
        }

        spreadsheetView.getFixedRows().clear();
        for (int rowIndex = 0; rowIndex < Math.min(renderState.frozenRowCount(), renderState.rowCount()); rowIndex++) {
            spreadsheetView.getFixedRows().add(rowIndex);
        }

        spreadsheetView.getFixedColumns().clear();
        for (int columnIndex = 0; columnIndex < Math.min(renderState.frozenColumnCount(), columns.size()); columnIndex++) {
            spreadsheetView.getFixedColumns().add(columns.get(columnIndex));
        }
    }

    private void enableGapClosingColumnScaling(SpreadsheetView spreadsheetView, BitSet hiddenColumns) {
        if (Boolean.TRUE.equals(spreadsheetView.getProperties().get(GAP_SCALING_ENABLED_KEY))) {
            return;
        }

        Map<SpreadsheetColumn, Double> originalWidths = new HashMap<>();
        spreadsheetView.getProperties().put(GAP_SCALING_ENABLED_KEY, true);

        Platform.runLater(() -> {
            List<SpreadsheetColumn> columns = spreadsheetView.getColumns();
            for (int columnIndex = 0; columnIndex < columns.size(); columnIndex++) {
                SpreadsheetColumn column = columns.get(columnIndex);
                column.setResizable(true);
                column.setMinWidth(MIN_RENDERED_COLUMN_WIDTH);
                column.setMaxWidth(Double.MAX_VALUE);
                if (!hiddenColumns.get(columnIndex)) {
                    originalWidths.put(column, column.getWidth());
                }
            }

            scheduleGapClosingPasses(spreadsheetView, originalWidths, hiddenColumns, GAP_CLOSING_PASSES);

            spreadsheetView.widthProperty().addListener((observable, oldValue, newValue) ->
                    scheduleGapClosingPasses(spreadsheetView, originalWidths, hiddenColumns, 3)
            );

            spreadsheetView.heightProperty().addListener((observable, oldValue, newValue) ->
                    scheduleGapClosingPasses(spreadsheetView, originalWidths, hiddenColumns, 3)
            );

            spreadsheetView.sceneProperty().addListener((observable, oldScene, newScene) -> {
                if (newScene != null) {
                    scheduleGapClosingPasses(spreadsheetView, originalWidths, hiddenColumns, GAP_CLOSING_PASSES);
                }
            });
        });
    }

    private void scheduleGapClosingPasses(
            SpreadsheetView spreadsheetView,
            Map<SpreadsheetColumn, Double> originalWidths,
            BitSet hiddenColumns,
            int remainingPasses
    ) {
        if (remainingPasses <= 0) {
            return;
        }

        Platform.runLater(() -> {
            fitColumnsToViewport(spreadsheetView, originalWidths, hiddenColumns);

            if (spreadsheetView.getScene() != null
                    && ((shouldForceWidthFit(spreadsheetView) && isHorizontalScrollBarVisible(spreadsheetView))
                    || Math.abs(calculateRightSideGap(spreadsheetView)) > 1.0)) {
                scheduleGapClosingPasses(spreadsheetView, originalWidths, hiddenColumns, remainingPasses - 1);
            }
        });
    }

    private void fitColumnsToViewport(
            SpreadsheetView spreadsheetView,
            Map<SpreadsheetColumn, Double> originalWidths,
            BitSet hiddenColumns
    ) {
        if (spreadsheetView.getColumns().isEmpty() || originalWidths.isEmpty()) {
            return;
        }

        spreadsheetView.applyCss();
        spreadsheetView.layout();

        double availableWidth = calculateAvailableColumnAreaWidth(spreadsheetView);
        if (availableWidth <= 0) {
            return;
        }

        List<SpreadsheetColumn> columns = spreadsheetView.getColumns();
        double totalBaseWidth = 0.0;
        for (int columnIndex = 0; columnIndex < columns.size(); columnIndex++) {
            SpreadsheetColumn column = columns.get(columnIndex);
            if (!hiddenColumns.get(columnIndex)) {
                totalBaseWidth += originalWidths.getOrDefault(column, column.getWidth());
            }
        }

        if (totalBaseWidth <= 0) {
            return;
        }

        boolean shouldForceWidthFit = shouldForceWidthFit(spreadsheetView);
        double scaleFactor = shouldForceWidthFit
                ? availableWidth / totalBaseWidth
                : Math.max(1.0, availableWidth / totalBaseWidth);

        for (int columnIndex = 0; columnIndex < columns.size(); columnIndex++) {
            SpreadsheetColumn column = columns.get(columnIndex);
            if (hiddenColumns.get(columnIndex)) {
                continue;
            }

            double baseWidth = originalWidths.getOrDefault(column, column.getWidth());
            column.setPrefWidth(Math.max(MIN_RENDERED_COLUMN_WIDTH, baseWidth * scaleFactor));
        }

        spreadsheetView.applyCss();
        spreadsheetView.layout();

        if (!isHorizontalScrollBarVisible(spreadsheetView)) {
            double gap = calculateRightSideGap(spreadsheetView);
            if (Math.abs(gap) > 0.75) {
                nudgeVisibleColumnsByGap(spreadsheetView, hiddenColumns, gap);
            }
        }
    }

    private boolean shouldForceWidthFit(SpreadsheetView spreadsheetView) {
        if (spreadsheetView.getScene() == null || spreadsheetView.getScene().getWindow() == null) {
            return false;
        }

        if (spreadsheetView.getScene().getWindow() instanceof Stage stage) {
            return stage.isFullScreen() || stage.isMaximized();
        }

        return false;
    }

    private void nudgeVisibleColumnsByGap(
            SpreadsheetView spreadsheetView,
            BitSet hiddenColumns,
            double gap
    ) {
        List<SpreadsheetColumn> columns = spreadsheetView.getColumns();
        double currentTotalWidth = 0.0;
        for (int columnIndex = 0; columnIndex < columns.size(); columnIndex++) {
            if (!hiddenColumns.get(columnIndex)) {
                currentTotalWidth += columns.get(columnIndex).getWidth();
            }
        }

        if (currentTotalWidth <= 0) {
            return;
        }

        for (int columnIndex = 0; columnIndex < columns.size(); columnIndex++) {
            if (hiddenColumns.get(columnIndex)) {
                continue;
            }

            SpreadsheetColumn column = columns.get(columnIndex);
            double proportion = column.getWidth() / currentTotalWidth;
            double adjustedWidth = column.getWidth() + (gap * proportion);
            column.setPrefWidth(Math.max(MIN_RENDERED_COLUMN_WIDTH, adjustedWidth));
        }
    }

    private double calculateAvailableColumnAreaWidth(SpreadsheetView spreadsheetView) {
        Node firstColumnHeader = findFirstColumnHeader(spreadsheetView);

        if (firstColumnHeader == null) {
            return 0.0;
        }

        Bounds firstColumnBounds = firstColumnHeader.localToScene(firstColumnHeader.getBoundsInLocal());
        double leftEdgeOfFirstColumn = firstColumnBounds.getMinX();

        Node verticalScrollBar = findVerticalScrollBar(spreadsheetView);
        if (verticalScrollBar != null) {
            Bounds scrollBarBounds = verticalScrollBar.localToScene(verticalScrollBar.getBoundsInLocal());
            return Math.max(0.0, scrollBarBounds.getMinX() - leftEdgeOfFirstColumn);
        }

        Bounds spreadsheetBounds = spreadsheetView.localToScene(spreadsheetView.getBoundsInLocal());
        return Math.max(0.0, spreadsheetBounds.getMaxX() - leftEdgeOfFirstColumn - 2.0);
    }

    private double calculateRightSideGap(SpreadsheetView spreadsheetView) {
        Node verticalScrollBar = findVerticalScrollBar(spreadsheetView);
        Node lastColumnHeader = findLastColumnHeader(spreadsheetView);

        if (verticalScrollBar == null || lastColumnHeader == null) {
            return 0.0;
        }

        Bounds scrollBarBounds = verticalScrollBar.localToScene(verticalScrollBar.getBoundsInLocal());
        Bounds lastColumnBounds = lastColumnHeader.localToScene(lastColumnHeader.getBoundsInLocal());

        double leftEdgeOfScrollBar = scrollBarBounds.getMinX();
        double rightEdgeOfLastColumn = lastColumnBounds.getMaxX();
        return leftEdgeOfScrollBar - rightEdgeOfLastColumn;
    }

    private boolean isHorizontalScrollBarVisible(SpreadsheetView spreadsheetView) {
        Set<Node> scrollBars = spreadsheetView.lookupAll(".scroll-bar");

        for (Node node : scrollBars) {
            if (node instanceof ScrollBar scrollBar
                    && scrollBar.getOrientation() == Orientation.HORIZONTAL
                    && scrollBar.isVisible()) {
                return true;
            }
        }

        return false;
    }

    private Node findVerticalScrollBar(SpreadsheetView spreadsheetView) {
        Set<Node> scrollBars = spreadsheetView.lookupAll(".scroll-bar");

        for (Node node : scrollBars) {
            if (node instanceof ScrollBar scrollBar
                    && scrollBar.getOrientation() == Orientation.VERTICAL
                    && scrollBar.isVisible()) {
                return scrollBar;
            }
        }

        return null;
    }

    private Node findFirstColumnHeader(SpreadsheetView spreadsheetView) {
        Set<Node> columnHeaders = spreadsheetView.lookupAll(".column-header");

        Node leftMostHeader = null;
        double leftMostX = Double.POSITIVE_INFINITY;

        for (Node node : columnHeaders) {
            Bounds bounds = node.localToScene(node.getBoundsInLocal());
            if (bounds.getWidth() <= 0) {
                continue;
            }

            if (bounds.getMinX() < leftMostX) {
                leftMostX = bounds.getMinX();
                leftMostHeader = node;
            }
        }

        return leftMostHeader;
    }

    private Node findLastColumnHeader(SpreadsheetView spreadsheetView) {
        Set<Node> columnHeaders = spreadsheetView.lookupAll(".column-header");

        Node rightMostHeader = null;
        double rightMostX = Double.NEGATIVE_INFINITY;

        for (Node node : columnHeaders) {
            Bounds bounds = node.localToScene(node.getBoundsInLocal());
            if (bounds.getWidth() <= 0) {
                continue;
            }

            if (bounds.getMaxX() > rightMostX) {
                rightMostX = bounds.getMaxX();
                rightMostHeader = node;
            }
        }

        return rightMostHeader;
    }

    private void refreshDisplayedCellStyle(
            SpreadsheetView spreadsheetView,
            Workbook workbook,
            Sheet sheet,
            List<CellRangeAddress> mergedRegions,
            int rowCount,
            int columnCount,
            int rowIndex,
            int columnIndex
    ) {
        if (rowIndex >= spreadsheetView.getGrid().getRowCount() || columnIndex >= spreadsheetView.getGrid().getColumnCount()) {
            return;
        }

        CellStyle cellStyle = getCellStyle(sheet, rowIndex, columnIndex);
        if (cellStyle == null) {
            cellStyle = workbook.getCellStyleAt(0);
        }

        CellRangeAddress mergedRegion = findMergedRegionAnchor(mergedRegions, rowIndex, columnIndex);
        String style = buildCellStyle(workbook, sheet, rowIndex, columnIndex, rowCount, columnCount, mergedRegion, cellStyle);
        spreadsheetView.getGrid().getRows().get(rowIndex).get(columnIndex).setStyle(style);
    }

    private CellRangeAddress findMergedRegionAnchor(
            List<CellRangeAddress> mergedRegions,
            int rowIndex,
            int columnIndex
    ) {
        for (CellRangeAddress mergedRegion : mergedRegions) {
            if (mergedRegion.getFirstRow() <= rowIndex
                    && mergedRegion.getLastRow() >= rowIndex
                    && mergedRegion.getFirstColumn() <= columnIndex
                    && mergedRegion.getLastColumn() >= columnIndex) {
                return mergedRegion;
            }
        }

        return null;
    }

    private String buildCellStyle(
            Workbook workbook,
            Sheet sheet,
            int rowIndex,
            int columnIndex,
            int rowCount,
            int columnCount,
            CellRangeAddress mergedRegion,
            CellStyle cellStyle
    ) {
        StringBuilder styleBuilder = new StringBuilder();
        appendAlignment(styleBuilder, cellStyle.getAlignment(), cellStyle.getVerticalAlignment());
        appendFill(styleBuilder, cellStyle, workbook);
        appendBorders(styleBuilder, workbook, sheet, rowIndex, columnIndex, rowCount, columnCount, mergedRegion, cellStyle);
        appendFont(styleBuilder, workbook, cellStyle);
        return styleBuilder.toString();
    }

    private void appendAlignment(StringBuilder styleBuilder, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment) {
        String vertical = switch (verticalAlignment) {
            case TOP, JUSTIFY, DISTRIBUTED -> "TOP";
            case BOTTOM -> "BOTTOM";
            default -> "CENTER";
        };

        String horizontal = switch (horizontalAlignment) {
            case CENTER, CENTER_SELECTION, FILL -> "CENTER";
            case RIGHT -> "RIGHT";
            default -> "LEFT";
        };

        String alignment = "CENTER".equals(vertical) && "CENTER".equals(horizontal)
                ? "CENTER"
                : vertical + '_' + horizontal;

        styleBuilder.append("-fx-alignment: ").append(alignment).append(';');
    }

    private void appendFill(StringBuilder styleBuilder, CellStyle cellStyle, Workbook workbook) {
        if (cellStyle.getFillPattern() == FillPatternType.NO_FILL) {
            return;
        }

        String fillColor = resolveFillColor(cellStyle, workbook);
        if (fillColor != null) {
            styleBuilder.append("-fx-background-color: ").append(fillColor).append(';');
            styleBuilder.append("-fx-background-insets: 0;");
            styleBuilder.append("-fx-background-radius: 0;");
        }
    }

    private void appendBorders(
            StringBuilder styleBuilder,
            Workbook workbook,
            Sheet sheet,
            int rowIndex,
            int columnIndex,
            int rowCount,
            int columnCount,
            CellRangeAddress mergedRegion,
            CellStyle cellStyle
    ) {
        BorderRenderAttributes topBorder;
        BorderRenderAttributes rightBorder;
        BorderRenderAttributes bottomBorder;
        BorderRenderAttributes leftBorder;

        if (mergedRegion != null) {
            topBorder = ownedMergedTopBorder(workbook, sheet, mergedRegion);
            rightBorder = ownedMergedRightBorder(workbook, sheet, mergedRegion, columnCount);
            bottomBorder = ownedMergedBottomBorder(workbook, sheet, mergedRegion, rowCount, columnCount);
            leftBorder = ownedMergedLeftBorder(workbook, sheet, mergedRegion, rowCount, columnCount);
        } else {
            topBorder = ownedTopBorder(workbook, sheet, rowIndex, columnIndex, cellStyle);
            rightBorder = ownedRightBorder(workbook, cellStyle, columnIndex, columnCount);
            bottomBorder = ownedBottomBorder(workbook, cellStyle, sheet, rowIndex, columnIndex, rowCount);
            leftBorder = ownedLeftBorder(workbook, sheet, rowIndex, columnIndex, cellStyle);
        }

        styleBuilder.append("-fx-border-width: ")
                .append(topBorder.width()).append(' ')
                .append(rightBorder.width()).append(' ')
                .append(bottomBorder.width()).append(' ')
                .append(leftBorder.width()).append(';');
        styleBuilder.append("-fx-border-style: ")
                .append(topBorder.style()).append(' ')
                .append(rightBorder.style()).append(' ')
                .append(bottomBorder.style()).append(' ')
                .append(leftBorder.style()).append(';');
        styleBuilder.append("-fx-border-color: ")
                .append(topBorder.color()).append(' ')
                .append(rightBorder.color()).append(' ')
                .append(bottomBorder.color()).append(' ')
                .append(leftBorder.color()).append(';');
        styleBuilder.append("-fx-border-insets: 0;");
        styleBuilder.append("-fx-border-radius: 0;");
    }

    private BorderRenderAttributes ownedMergedTopBorder(
            Workbook workbook,
            Sheet sheet,
            CellRangeAddress mergedRegion
    ) {
        if (mergedRegion.getFirstRow() == 0) {
            return strongestRegionEdgeBorder(workbook, sheet, mergedRegion, BorderSide.TOP);
        }

        return noBorderAttributes();
    }

    private BorderRenderAttributes ownedMergedRightBorder(
            Workbook workbook,
            Sheet sheet,
            CellRangeAddress mergedRegion,
            int columnCount
    ) {
        if (mergedRegion.getLastColumn() == columnCount - 1) {
            return strongestRegionEdgeBorder(workbook, sheet, mergedRegion, BorderSide.RIGHT);
        }

        return noBorderAttributes();
    }

    private BorderRenderAttributes ownedMergedBottomBorder(
            Workbook workbook,
            Sheet sheet,
            CellRangeAddress mergedRegion,
            int rowCount,
            int columnCount
    ) {
        BorderRenderAttributes regionEdge = strongestRegionEdgeBorder(workbook, sheet, mergedRegion, BorderSide.BOTTOM);
        BorderRenderAttributes adjacentEdge = strongestAdjacentBoundaryBorder(
                workbook,
                sheet,
                mergedRegion,
                BorderSide.BOTTOM,
                rowCount,
                columnCount
        );
        return mergeBorderAttributes(regionEdge, adjacentEdge);
    }

    private BorderRenderAttributes ownedMergedLeftBorder(
            Workbook workbook,
            Sheet sheet,
            CellRangeAddress mergedRegion,
            int rowCount,
            int columnCount
    ) {
        BorderRenderAttributes adjacentEdge = strongestAdjacentBoundaryBorder(
                workbook,
                sheet,
                mergedRegion,
                BorderSide.LEFT,
                rowCount,
                columnCount
        );
        BorderRenderAttributes regionEdge = strongestRegionEdgeBorder(workbook, sheet, mergedRegion, BorderSide.LEFT);
        return mergeBorderAttributes(adjacentEdge, regionEdge);
    }

    private BorderRenderAttributes strongestRegionEdgeBorder(
            Workbook workbook,
            Sheet sheet,
            CellRangeAddress mergedRegion,
            BorderSide borderSide
    ) {
        BorderRenderAttributes strongestBorder = baseGridlineAttributes();

        switch (borderSide) {
            case TOP -> {
                int rowIndex = mergedRegion.getFirstRow();
                for (int columnIndex = mergedRegion.getFirstColumn(); columnIndex <= mergedRegion.getLastColumn(); columnIndex++) {
                    strongestBorder = mergeBorderAttributes(
                            strongestBorder,
                            resolveBorderAttributes(workbook, getCellStyle(sheet, rowIndex, columnIndex), BorderSide.TOP)
                    );
                }
            }
            case RIGHT -> {
                int columnIndex = mergedRegion.getLastColumn();
                for (int rowIndex = mergedRegion.getFirstRow(); rowIndex <= mergedRegion.getLastRow(); rowIndex++) {
                    strongestBorder = mergeBorderAttributes(
                            strongestBorder,
                            resolveBorderAttributes(workbook, getCellStyle(sheet, rowIndex, columnIndex), BorderSide.RIGHT)
                    );
                }
            }
            case BOTTOM -> {
                int rowIndex = mergedRegion.getLastRow();
                for (int columnIndex = mergedRegion.getFirstColumn(); columnIndex <= mergedRegion.getLastColumn(); columnIndex++) {
                    strongestBorder = mergeBorderAttributes(
                            strongestBorder,
                            resolveBorderAttributes(workbook, getCellStyle(sheet, rowIndex, columnIndex), BorderSide.BOTTOM)
                    );
                }
            }
            case LEFT -> {
                int columnIndex = mergedRegion.getFirstColumn();
                for (int rowIndex = mergedRegion.getFirstRow(); rowIndex <= mergedRegion.getLastRow(); rowIndex++) {
                    strongestBorder = mergeBorderAttributes(
                            strongestBorder,
                            resolveBorderAttributes(workbook, getCellStyle(sheet, rowIndex, columnIndex), BorderSide.LEFT)
                    );
                }
            }
        }

        return strongestBorder;
    }

    private BorderRenderAttributes strongestAdjacentBoundaryBorder(
            Workbook workbook,
            Sheet sheet,
            CellRangeAddress mergedRegion,
            BorderSide borderSide,
            int rowCount,
            int columnCount
    ) {
        BorderRenderAttributes strongestBorder = noBorderAttributes();

        switch (borderSide) {
            case TOP -> {
                int adjacentRowIndex = mergedRegion.getFirstRow() - 1;
                if (adjacentRowIndex < 0) {
                    return strongestBorder;
                }
                for (int columnIndex = mergedRegion.getFirstColumn(); columnIndex <= mergedRegion.getLastColumn(); columnIndex++) {
                    strongestBorder = mergeBorderAttributes(
                            strongestBorder,
                            resolveBorderAttributes(workbook, getCellStyle(sheet, adjacentRowIndex, columnIndex), BorderSide.BOTTOM)
                    );
                }
            }
            case RIGHT -> {
                int adjacentColumnIndex = mergedRegion.getLastColumn() + 1;
                if (adjacentColumnIndex >= columnCount) {
                    return strongestBorder;
                }
                for (int rowIndex = mergedRegion.getFirstRow(); rowIndex <= mergedRegion.getLastRow(); rowIndex++) {
                    strongestBorder = mergeBorderAttributes(
                            strongestBorder,
                            resolveBorderAttributes(workbook, getCellStyle(sheet, rowIndex, adjacentColumnIndex), BorderSide.LEFT)
                    );
                }
            }
            case BOTTOM -> {
                int adjacentRowIndex = mergedRegion.getLastRow() + 1;
                if (adjacentRowIndex >= rowCount) {
                    return strongestBorder;
                }
                for (int columnIndex = mergedRegion.getFirstColumn(); columnIndex <= mergedRegion.getLastColumn(); columnIndex++) {
                    strongestBorder = mergeBorderAttributes(
                            strongestBorder,
                            resolveBorderAttributes(workbook, getCellStyle(sheet, adjacentRowIndex, columnIndex), BorderSide.TOP)
                    );
                }
            }
            case LEFT -> {
                int adjacentColumnIndex = mergedRegion.getFirstColumn() - 1;
                if (adjacentColumnIndex < 0) {
                    return strongestBorder;
                }
                for (int rowIndex = mergedRegion.getFirstRow(); rowIndex <= mergedRegion.getLastRow(); rowIndex++) {
                    strongestBorder = mergeBorderAttributes(
                            strongestBorder,
                            resolveBorderAttributes(workbook, getCellStyle(sheet, rowIndex, adjacentColumnIndex), BorderSide.RIGHT)
                    );
                }
            }
        }

        return strongestBorder;
    }

    private BorderRenderAttributes ownedTopBorder(
            Workbook workbook,
            Sheet sheet,
            int rowIndex,
            int columnIndex,
            CellStyle currentStyle
    ) {
        if (rowIndex == 0) {
            return resolveBorderAttributes(workbook, currentStyle, BorderSide.TOP);
        }

        return noBorderAttributes();
    }

    private BorderRenderAttributes ownedRightBorder(
            Workbook workbook,
            CellStyle currentStyle,
            int columnIndex,
            int columnCount
    ) {
        if (columnIndex == columnCount - 1) {
            return resolveBorderAttributes(workbook, currentStyle, BorderSide.RIGHT);
        }

        return noBorderAttributes();
    }

    private BorderRenderAttributes ownedBottomBorder(
            Workbook workbook,
            CellStyle currentStyle,
            Sheet sheet,
            int rowIndex,
            int columnIndex,
            int rowCount
    ) {
        if (rowIndex == rowCount - 1) {
            return resolveBorderAttributes(workbook, currentStyle, BorderSide.BOTTOM);
        }

        CellStyle belowStyle = getCellStyle(sheet, rowIndex + 1, columnIndex);
        return mergeBorderAttributes(
                resolveBorderAttributes(workbook, currentStyle, BorderSide.BOTTOM),
                resolveBorderAttributes(workbook, belowStyle, BorderSide.TOP)
        );
    }

    private BorderRenderAttributes ownedLeftBorder(
            Workbook workbook,
            Sheet sheet,
            int rowIndex,
            int columnIndex,
            CellStyle currentStyle
    ) {
        if (columnIndex == 0) {
            return resolveBorderAttributes(workbook, currentStyle, BorderSide.LEFT);
        }

        CellStyle leftStyle = getCellStyle(sheet, rowIndex, columnIndex - 1);
        return mergeBorderAttributes(
                resolveBorderAttributes(workbook, leftStyle, BorderSide.RIGHT),
                resolveBorderAttributes(workbook, currentStyle, BorderSide.LEFT)
        );
    }

    private BorderRenderAttributes resolveBorderAttributes(Workbook workbook, CellStyle cellStyle, BorderSide borderSide) {
        if (cellStyle == null) {
            return baseGridlineAttributes();
        }

        BorderStyle borderStyle = switch (borderSide) {
            case TOP -> cellStyle.getBorderTop();
            case RIGHT -> cellStyle.getBorderRight();
            case BOTTOM -> cellStyle.getBorderBottom();
            case LEFT -> cellStyle.getBorderLeft();
        };

        if (borderStyle == BorderStyle.NONE) {
            return baseGridlineAttributes();
        }

        return new BorderRenderAttributes(
                borderWidth(borderStyle),
                borderCssStyle(borderStyle),
                resolveBorderColor(workbook, cellStyle, borderSide),
                borderWeightClass(borderStyle)
        );
    }

    private BorderRenderAttributes mergeBorderAttributes(
            BorderRenderAttributes firstBorder,
            BorderRenderAttributes secondBorder
    ) {
        int firstPriority = borderPriority(firstBorder.weightClass());
        int secondPriority = borderPriority(secondBorder.weightClass());

        if (secondPriority > firstPriority) {
            return secondBorder;
        }
        if (firstPriority > secondPriority) {
            return firstBorder;
        }

        if (!"solid".equals(secondBorder.style())) {
            return secondBorder;
        }
        return firstBorder;
    }

    private int borderPriority(BorderWeightClass weightClass) {
        return switch (weightClass) {
            case NONE -> 0;
            case HAIR -> 1;
            case THIN -> 2;
            case MEDIUM -> 3;
            case THICK -> 4;
        };
    }

    private BorderRenderAttributes baseGridlineAttributes() {
        return new BorderRenderAttributes(BASE_GRIDLINE_WIDTH, "solid", BASE_GRIDLINE_COLOR, BorderWeightClass.NONE);
    }

    private BorderRenderAttributes noBorderAttributes() {
        return new BorderRenderAttributes("0", "none", "transparent", BorderWeightClass.NONE);
    }

    private CellStyle getCellStyle(Sheet sheet, int rowIndex, int columnIndex) {
        if (rowIndex < 0 || columnIndex < 0) {
            return null;
        }

        Cell cell = getCell(sheet, rowIndex, columnIndex);
        return cell == null ? null : cell.getCellStyle();
    }

    private Cell getCell(Sheet sheet, int rowIndex, int columnIndex) {
        if (rowIndex < 0 || columnIndex < 0) {
            return null;
        }

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return null;
        }

        return row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
    }

    private void appendFont(StringBuilder styleBuilder, Workbook workbook, CellStyle cellStyle) {
        Font font = workbook.getFontAt(cellStyle.getFontIndex());
        styleBuilder.append("-fx-font-family: '").append(escapeCss(font.getFontName())).append("';");
        styleBuilder.append("-fx-font-size: ").append(fontPointsToPixels(font.getFontHeightInPoints())).append("px;");
        if (font.getBold()) {
            styleBuilder.append("-fx-font-weight: bold;");
        }
        if (font.getItalic()) {
            styleBuilder.append("-fx-font-style: italic;");
        }
        if (font.getUnderline() != Font.U_NONE) {
            styleBuilder.append("-fx-underline: true;");
        }

        String textColor = resolveFontColor(font, workbook);
        if (textColor != null) {
            styleBuilder.append("-fx-text-fill: ").append(textColor).append(';');
        }
    }

    private String resolveFillColor(CellStyle cellStyle, Workbook workbook) {
        if (cellStyle instanceof XSSFCellStyle xssfCellStyle) {
            String color = toHex(xssfCellStyle.getFillForegroundColorColor());
            if (color == null) {
                color = toHex(xssfCellStyle.getFillForegroundXSSFColor());
            }
            if (color == null) {
                color = indexedOrPaletteColorToHex(xssfCellStyle.getFillForegroundColor(), workbook);
            }
            if (color == null) {
                color = toHex(xssfCellStyle.getFillBackgroundColorColor());
            }
            if (color == null) {
                color = toHex(xssfCellStyle.getFillBackgroundXSSFColor());
            }
            if (color == null) {
                color = indexedOrPaletteColorToHex(xssfCellStyle.getFillBackgroundColor(), workbook);
            }
            return color;
        }
        return indexedOrPaletteColorToHex(cellStyle.getFillForegroundColor(), workbook);
    }

    private String resolveBorderColor(Workbook workbook, CellStyle cellStyle, BorderSide borderSide) {
        if (cellStyle instanceof XSSFCellStyle xssfCellStyle) {
            String xssfColor = switch (borderSide) {
                case TOP -> toHex(xssfCellStyle.getTopBorderXSSFColor());
                case RIGHT -> toHex(xssfCellStyle.getRightBorderXSSFColor());
                case BOTTOM -> toHex(xssfCellStyle.getBottomBorderXSSFColor());
                case LEFT -> toHex(xssfCellStyle.getLeftBorderXSSFColor());
            };
            if (xssfColor != null) {
                return xssfColor;
            }
        }

        short colorIndex = switch (borderSide) {
            case TOP -> cellStyle.getTopBorderColor();
            case RIGHT -> cellStyle.getRightBorderColor();
            case BOTTOM -> cellStyle.getBottomBorderColor();
            case LEFT -> cellStyle.getLeftBorderColor();
        };
        return orDefault(indexedOrPaletteColorToHex(colorIndex, workbook));
    }

    private String resolveFontColor(Font font, Workbook workbook) {
        if (font instanceof XSSFFont xssfFont) {
            String color = toHex(xssfFont.getXSSFColor());
            if (color != null) {
                return color;
            }
        }
        return indexedOrPaletteColorToHex(font.getColor(), workbook);
    }

    private String indexedOrPaletteColorToHex(short colorIndex, Workbook workbook) {
        if (workbook instanceof HSSFWorkbook hssfWorkbook) {
            HSSFPalette palette = hssfWorkbook.getCustomPalette();
            HSSFColor color = palette.getColor(colorIndex);
            String paletteColor = toHex(color);
            if (paletteColor != null) {
                return paletteColor;
            }
        }

        return indexedColorToHex(colorIndex);
    }

    private String indexedColorToHex(short colorIndex) {
        IndexedColors indexedColor = IndexedColors.fromInt(colorIndex);
        if (indexedColor == null || indexedColor == IndexedColors.AUTOMATIC) {
            return null;
        }

        return switch (indexedColor) {
            case BLACK, BLACK1 -> "#000000";
            case WHITE, WHITE1 -> "#FFFFFF";
            case RED, RED1, DARK_RED -> "#C00000";
            case BRIGHT_GREEN, BRIGHT_GREEN1, GREEN -> "#00B050";
            case BLUE, BLUE1, DARK_BLUE, ROYAL_BLUE -> "#4472C4";
            case YELLOW, YELLOW1, GOLD, LIGHT_YELLOW -> "#FFD966";
            case GREY_25_PERCENT -> "#D9D9D9";
            case GREY_40_PERCENT -> "#BFBFBF";
            case GREY_50_PERCENT -> "#7F7F7F";
            case GREY_80_PERCENT -> "#404040";
            case LIGHT_GREEN -> "#E2F0D9";
            case LIGHT_ORANGE, ORANGE -> "#F4B183";
            case LIGHT_BLUE, PALE_BLUE, SKY_BLUE -> "#DDEBF7";
            default -> "#000000";
        };
    }

    private String toHex(org.apache.poi.ss.usermodel.Color color) {
        if (color == null) {
            return null;
        }
        if (color instanceof XSSFColor xssfColor) {
            String argbHex = xssfColor.getARGBHex();
            if (argbHex != null && !argbHex.isBlank()) {
                return hexStringToCss(argbHex);
            }

            byte[] rgb = xssfColor.getRGB();
            if (rgb != null) {
                return rgbToHex(rgb, 0);
            }

            byte[] tintedRgb = xssfColor.getRGBWithTint();
            if (tintedRgb != null) {
                return rgbToHex(tintedRgb, 0);
            }

            byte[] argb = xssfColor.getARGB();
            if (argb != null) {
                return rgbToHex(argb, argb.length == 4 ? 1 : 0);
            }

            if (xssfColor.isIndexed()) {
                return indexedColorToHex(xssfColor.getIndexed());
            }
            return null;
        }
        if (color instanceof HSSFColor hssfColor) {
            short[] triplet = hssfColor.getTriplet();
            return tripletToHex(triplet);
        }
        return null;
    }

    private String rgbToHex(byte[] rgb, int offset) {
        if (rgb == null || rgb.length - offset < 3) {
            return null;
        }
        return "#%02X%02X%02X".formatted(
                rgb[offset] & 0xFF,
                rgb[offset + 1] & 0xFF,
                rgb[offset + 2] & 0xFF
        );
    }

    private String tripletToHex(short[] triplet) {
        if (triplet == null || triplet.length < 3) {
            return null;
        }
        return "#%02X%02X%02X".formatted(triplet[0], triplet[1], triplet[2]);
    }

    private String hexStringToCss(String argbHex) {
        String normalizedHex = argbHex.trim();
        if (normalizedHex.length() == 8) {
            normalizedHex = normalizedHex.substring(2);
        }
        if (normalizedHex.length() != 6) {
            return null;
        }
        return '#' + normalizedHex.toUpperCase();
    }

    private String borderWidth(BorderStyle borderStyle) {
        return switch (borderStyle) {
            case NONE -> BASE_GRIDLINE_WIDTH;
            case HAIR -> "0.2";
            case MEDIUM, MEDIUM_DASHED, MEDIUM_DASH_DOT, MEDIUM_DASH_DOT_DOT -> "0.75";
            case DOUBLE -> "0.95";
            case THICK -> "1.1";
            default -> "0.42";
        };
    }

    private BorderWeightClass borderWeightClass(BorderStyle borderStyle) {
        return switch (borderStyle) {
            case NONE -> BorderWeightClass.NONE;
            case HAIR -> BorderWeightClass.HAIR;
            case MEDIUM, MEDIUM_DASHED, MEDIUM_DASH_DOT, MEDIUM_DASH_DOT_DOT, DOUBLE -> BorderWeightClass.MEDIUM;
            case THICK -> BorderWeightClass.THICK;
            default -> BorderWeightClass.THIN;
        };
    }

    private String borderCssStyle(BorderStyle borderStyle) {
        return switch (borderStyle) {
            case NONE -> "none";
            case DASHED, MEDIUM_DASHED, DASH_DOT, DASH_DOT_DOT, MEDIUM_DASH_DOT, MEDIUM_DASH_DOT_DOT, SLANTED_DASH_DOT -> "dashed";
            case DOTTED -> "dotted";
            default -> "solid";
        };
    }

    private String orDefault(String color) {
        return color == null ? "#000000" : color;
    }

    private double rowHeightPointsToPixels(double points) {
        return Math.max(1, points * 96d / 72d);
    }

    private double fontPointsToPixels(double points) {
        return Math.max(1, points * 96d / 72d);
    }

    private String escapeCss(String value) {
        return value.replace("'", "\\'");
    }

    private enum BorderSide {
        TOP,
        RIGHT,
        BOTTOM,
        LEFT
    }

    private enum BorderWeightClass {
        NONE,
        HAIR,
        THIN,
        MEDIUM,
        THICK
    }

    private record BorderRenderAttributes(
            String width,
            String style,
            String color,
            BorderWeightClass weightClass
    ) {
    }

    private record SpreadsheetRenderState(
            ObservableList<ObservableList<SpreadsheetCell>> rows,
            int rowCount,
            int columnCount,
            Map<Integer, Double> rowHeights,
            double defaultRowHeight,
            List<String> rowHeaders,
            List<String> columnHeaders,
            List<Double> columnWidths,
            BitSet hiddenRows,
            BitSet hiddenColumns,
            List<CellRangeAddress> mergedRegions,
            int frozenRowCount,
            int frozenColumnCount
    ) {
    }
}
