package org.example.service;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.stream.Stream;

public class CrewMemberDirectoryService {

    private static final int CREW_NAME_COLUMN_INDEX = 1;
    private static final int INVENTORY_NAME_COLUMN_INDEX = 2;
    private static final int FIRST_NAME_ROW_INDEX = 2;
    private static final String CREW_WORKBOOK_PREFIX = "crew";
    private static final String INVENTORY_WORKBOOK_PREFIX = "inventory";

    private final Path dataDirectory;

    public CrewMemberDirectoryService() {
        this(ApplicationDataDirectoryResolver.resolveDataDirectory());
    }

    CrewMemberDirectoryService(Path dataDirectory) {
        this.dataDirectory = dataDirectory;
    }

    public List<String> loadCrewMemberNames() throws IOException {
        return loadNamesFromWorkbook(locateCrewWorkbook(), "Crew", CREW_NAME_COLUMN_INDEX);
    }

    public List<String> loadMateNames() throws IOException {
        return loadNamesFromWorkbook(locateInventoryWorkbook(), "Inventory", INVENTORY_NAME_COLUMN_INDEX);
    }

    public List<String> loadCrewMemberNamesExcludingMates() throws IOException {
        List<String> crewMemberNames = loadCrewMemberNames();
        Set<String> mateNames = new HashSet<>();
        for (String mateName : loadMateNames()) {
            mateNames.add(normalizeName(mateName));
        }

        List<String> filteredCrewMemberNames = new ArrayList<>();
        for (String crewMemberName : crewMemberNames) {
            if (!mateNames.contains(normalizeName(crewMemberName))) {
                filteredCrewMemberNames.add(crewMemberName);
            }
        }

        return filteredCrewMemberNames;
    }

    public Path locateCrewWorkbook() throws IOException {
        return locateWorkbook(CREW_WORKBOOK_PREFIX, "Crew");
    }

    public Path locateInventoryWorkbook() throws IOException {
        return locateWorkbook(INVENTORY_WORKBOOK_PREFIX, "Inventory");
    }

    private List<String> loadNamesFromWorkbook(Path workbookPath, String workbookLabel, int nameColumnIndex) throws IOException {
        try (InputStream inputStream = Files.newInputStream(workbookPath);
             Workbook workbook = WorkbookFactory.create(inputStream)) {
            if (workbook.getNumberOfSheets() == 0) {
                throw new IOException("The %s workbook does not contain any sheets.".formatted(workbookLabel));
            }

            return readNames(workbook.getSheetAt(0), new DataFormatter(Locale.US), nameColumnIndex);
        }
    }

    private Path locateWorkbook(String workbookPrefix, String workbookLabel) throws IOException {
        if (!Files.isDirectory(dataDirectory)) {
            throw new IOException("The data folder does not exist: %s".formatted(dataDirectory));
        }

        try (Stream<Path> fileStream = Files.list(dataDirectory)) {
            return fileStream
                    .filter(Files::isRegularFile)
                    .filter(filePath -> isWorkbookMatch(filePath, workbookPrefix))
                    .max(Comparator.comparing(this::lastModifiedTime))
                    .orElseThrow(() -> new IOException(
                            "No %s workbook was found in %s.".formatted(workbookLabel, dataDirectory)
                    ));
        }
    }

    private List<String> readNames(Sheet sheet, DataFormatter dataFormatter, int nameColumnIndex) {
        List<String> names = new ArrayList<>();
        for (int rowIndex = FIRST_NAME_ROW_INDEX; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }

            String name = dataFormatter.formatCellValue(
                    row.getCell(nameColumnIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)
            ).trim();
            if (!name.isBlank()) {
                names.add(name);
            }
        }

        return names;
    }

    private boolean isWorkbookMatch(Path filePath, String workbookPrefix) {
        String fileName = filePath.getFileName().toString().toLowerCase(Locale.ROOT);
        return fileName.startsWith(workbookPrefix)
                && (fileName.endsWith(".xlsx") || fileName.endsWith(".xlsm") || fileName.endsWith(".xls"));
    }

    private String normalizeName(String name) {
        return name == null ? "" : name.trim().toLowerCase(Locale.ROOT);
    }

    private FileTime lastModifiedTime(Path filePath) {
        try {
            return Files.getLastModifiedTime(filePath);
        } catch (IOException exception) {
            return FileTime.fromMillis(Long.MIN_VALUE);
        }
    }
}
