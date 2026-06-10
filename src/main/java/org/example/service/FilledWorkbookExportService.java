package org.example.service;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

public class FilledWorkbookExportService {

    public Path exportSheetToNewWorkbook(Workbook sourceWorkbook, int sheetIndex, Path outputWorkbook) throws IOException {
        if (sourceWorkbook == null) {
            throw new IllegalArgumentException("A source workbook is required.");
        }
        if (sheetIndex < 0 || sheetIndex >= sourceWorkbook.getNumberOfSheets()) {
            throw new IllegalArgumentException("The selected sheet could not be found.");
        }

        Path parent = outputWorkbook.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        Sheet sourceSheet = sourceWorkbook.getSheetAt(sheetIndex);
        try (Workbook targetWorkbook = createTargetWorkbook(sourceWorkbook);
             OutputStream outputStream = Files.newOutputStream(outputWorkbook)) {
            PoiSheetToNewWorkbook.copySheetToNewWorkbook(
                    sourceSheet,
                    targetWorkbook,
                    sourceSheet.getSheetName()
            );
            targetWorkbook.setActiveSheet(0);
            targetWorkbook.setSelectedTab(0);
            targetWorkbook.write(outputStream);
        }

        return outputWorkbook;
    }

    public String resolveExportExtension(Workbook sourceWorkbook) {
        if (sourceWorkbook instanceof HSSFWorkbook) {
            return ".xls";
        }
        return ".xlsx";
    }

    private Workbook createTargetWorkbook(Workbook sourceWorkbook) {
        if (sourceWorkbook instanceof HSSFWorkbook) {
            return new HSSFWorkbook();
        }
        return new XSSFWorkbook();
    }
}
