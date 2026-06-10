package org.example.model;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

public class RegisterSheetState {

    private final Map<CellCoordinate, CellStyle> baselineStyles = new HashMap<>();
    private final Set<RegisterBlockKey> manualAssignments = new HashSet<>();
    private final Set<RegisterBlockKey> autoAssignments = new HashSet<>();
    private final Set<RegisterBlockKey> blockedForNextRun = new HashSet<>();
    private final Map<Integer, CrewRowXFillSnapshot> crewRowXFillSnapshots = new HashMap<>();

    public void rememberBaselineStyle(int rowIndex, int columnIndex, CellStyle cellStyle) {
        baselineStyles.putIfAbsent(new CellCoordinate(rowIndex, columnIndex), cellStyle);
    }

    public CellStyle getBaselineStyle(int rowIndex, int columnIndex) {
        return baselineStyles.get(new CellCoordinate(rowIndex, columnIndex));
    }

    public Set<RegisterBlockKey> getManualAssignments() {
        return manualAssignments;
    }

    public Set<RegisterBlockKey> getAutoAssignments() {
        return autoAssignments;
    }

    public Set<RegisterBlockKey> getBlockedForNextRun() {
        return blockedForNextRun;
    }

    public boolean isAssigned(RegisterBlockKey registerBlockKey) {
        return manualAssignments.contains(registerBlockKey) || autoAssignments.contains(registerBlockKey);
    }

    public CrewRowXFillSnapshot getCrewRowXFillSnapshot(int rowIndex) {
        return crewRowXFillSnapshots.get(rowIndex);
    }

    public void rememberCrewRowXFillSnapshot(int rowIndex, CrewRowXFillSnapshot snapshot) {
        crewRowXFillSnapshots.put(rowIndex, snapshot);
    }

    public CrewRowXFillSnapshot removeCrewRowXFillSnapshot(int rowIndex) {
        return crewRowXFillSnapshots.remove(rowIndex);
    }

    public Map<Integer, CrewRowXFillSnapshot> getCrewRowXFillSnapshots() {
        return crewRowXFillSnapshots;
    }

    public record CellCoordinate(int rowIndex, int columnIndex) {
    }

    public record CrewRowXFillSnapshot(
            Map<CellCoordinate, CellSnapshot> cellSnapshots,
            Set<RegisterBlockKey> manualAssignments,
            Set<RegisterBlockKey> autoAssignments
    ) {
    }

    public record CellSnapshot(
            boolean missing,
            CellStyle cellStyle,
            CellType cellType,
            String stringValue,
            Double numericValue,
            Boolean booleanValue,
            String formula,
            Byte errorValue
    ) {
    }
}
