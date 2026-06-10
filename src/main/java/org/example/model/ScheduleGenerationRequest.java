package org.example.model;

import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class ScheduleGenerationRequest {

    private final Path sourceWorkbook;
    private final String sourceSheetName;
    private final Path outputWorkbook;
    private final LocalDate scheduleDate;
    private final String storeName;
    private final String notes;
    private final List<String> taskNames;

    public ScheduleGenerationRequest(
            Path sourceWorkbook,
            String sourceSheetName,
            Path outputWorkbook,
            LocalDate scheduleDate,
            String storeName,
            String notes,
            List<String> taskNames
    ) {
        this.sourceWorkbook = sourceWorkbook;
        this.sourceSheetName = sourceSheetName;
        this.outputWorkbook = outputWorkbook;
        this.scheduleDate = scheduleDate;
        this.storeName = storeName;
        this.notes = notes;
        this.taskNames = Collections.unmodifiableList(new ArrayList<>(taskNames));
    }

    public Path getSourceWorkbook() {
        return sourceWorkbook;
    }

    public String getSourceSheetName() {
        return sourceSheetName;
    }

    public Path getOutputWorkbook() {
        return outputWorkbook;
    }

    public LocalDate getScheduleDate() {
        return scheduleDate;
    }

    public String getStoreName() {
        return storeName;
    }

    public String getNotes() {
        return notes;
    }

    public List<String> getTaskNames() {
        return taskNames;
    }
}
