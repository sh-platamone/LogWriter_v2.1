package org.example.model;

import java.nio.file.Path;

public class ScheduleGenerationResult {

    private final Path outputWorkbook;
    private final String generatedSheetName;
    private final int taskCount;

    public ScheduleGenerationResult(Path outputWorkbook, String generatedSheetName, int taskCount) {
        this.outputWorkbook = outputWorkbook;
        this.generatedSheetName = generatedSheetName;
        this.taskCount = taskCount;
    }

    public Path getOutputWorkbook() {
        return outputWorkbook;
    }

    public String getGeneratedSheetName() {
        return generatedSheetName;
    }

    public int getTaskCount() {
        return taskCount;
    }
}
