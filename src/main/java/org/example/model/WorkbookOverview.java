package org.example.model;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class WorkbookOverview {

    private final String workbookName;
    private final int sheetCount;
    private final String activeSheetName;
    private final List<SheetSummary> sheetSummaries;

    public WorkbookOverview(String workbookName, int sheetCount, String activeSheetName, List<SheetSummary> sheetSummaries) {
        this.workbookName = workbookName;
        this.sheetCount = sheetCount;
        this.activeSheetName = activeSheetName;
        this.sheetSummaries = Collections.unmodifiableList(new ArrayList<>(sheetSummaries));
    }

    public String getWorkbookName() {
        return workbookName;
    }

    public int getSheetCount() {
        return sheetCount;
    }

    public String getActiveSheetName() {
        return activeSheetName;
    }

    public List<SheetSummary> getSheetSummaries() {
        return sheetSummaries;
    }
}
