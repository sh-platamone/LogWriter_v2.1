package org.example.service;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.example.model.DaySheetOption;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

public class WorkbookInspectorService {

    private static final List<String> DAY_LABELS = List.of(
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
            "Sunday"
    );

    public List<DaySheetOption> loadDaySheetOptions(Path workbookPath) throws IOException {
        try (InputStream inputStream = Files.newInputStream(workbookPath);
             Workbook workbook = WorkbookFactory.create(inputStream)) {

            return loadDaySheetOptions(workbook);
        }
    }

    public List<DaySheetOption> loadDaySheetOptions(Workbook workbook) {
        if (workbook.getNumberOfSheets() < DAY_LABELS.size()) {
            throw new IllegalArgumentException(
                    "The workbook needs at least 7 sheets so Monday through Sunday can map to indices 0 through 6."
            );
        }

        List<DaySheetOption> daySheetOptions = new ArrayList<>();
        for (int index = 0; index < DAY_LABELS.size(); index++) {
            daySheetOptions.add(
                    new DaySheetOption(
                            DAY_LABELS.get(index),
                            index,
                            workbook.getSheetAt(index).getSheetName()
                    )
            );
        }

        return daySheetOptions;
    }
}
