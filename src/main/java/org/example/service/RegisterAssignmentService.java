package org.example.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.example.model.RegisterBlockKey;
import org.example.model.RegisterSheetState.CellCoordinate;
import org.example.model.RegisterSheetState.CellSnapshot;
import org.example.model.RegisterSheetState.CrewRowXFillSnapshot;
import org.example.model.RegisterSheetState;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.IdentityHashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.HashSet;
import java.util.Set;

public class RegisterAssignmentService {

    private static final short REGISTER_FILL_COLOR_INDEX = IndexedColors.YELLOW.getIndex();
    private static final String LEADERSHIP_TEAM_MARKER = "Leadership Team";

    private final Map<Workbook, Map<Short, CellStyle>> yellowStyleCacheByWorkbook = new IdentityHashMap<>();

    public RegisterSheetState createSheetState(Workbook workbook, Sheet sheet) {
        RegisterSheetState registerSheetState = new RegisterSheetState();
        captureBaselineStyles(workbook, sheet, registerSheetState);
        return registerSheetState;
    }

    public boolean toggleManualAssignment(
            Workbook workbook,
            Sheet sheet,
            RegisterSheetState registerSheetState,
            int rowIndex,
            int columnIndex
    ) {
        if (!RegisterGridLayout.isHourGridCoordinate(rowIndex, columnIndex)) {
            return false;
        }

        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        int leadershipTeamRowIndex = findLeadershipTeamRowIndex(sheet, formatter, evaluator);
        if (!isCrewRow(sheet, rowIndex, leadershipTeamRowIndex, formatter, evaluator)) {
            return false;
        }

        int firstColumnIndex = RegisterGridLayout.normalizeHourBlockStartColumn(columnIndex);
        RegisterBlockKey registerBlockKey = new RegisterBlockKey(rowIndex, firstColumnIndex);

        if (registerSheetState.isAssigned(registerBlockKey)) {
            restoreRegisterBlock(sheet, workbook, registerSheetState, registerBlockKey);
            registerSheetState.getManualAssignments().remove(registerBlockKey);
            registerSheetState.getAutoAssignments().remove(registerBlockKey);
            registerSheetState.getBlockedForNextRun().add(registerBlockKey);
            return true;
        }

        applyRegisterBlockFill(sheet, workbook, registerSheetState, registerBlockKey);
        registerSheetState.getManualAssignments().add(registerBlockKey);
        registerSheetState.getBlockedForNextRun().remove(registerBlockKey);
        return true;
    }

    public void applyAutomaticAssignments(
            Workbook workbook,
            Sheet sheet,
            RegisterSheetState registerSheetState,
            int requestedStartHour,
            int requestedEndHour,
            int targetRegisterCount,
            Map<String, Integer> floorPrioritiesByName
    ) {
        removeStoredRegisterAssignmentsFromCrewRowSnapshots(workbook, registerSheetState);
        clearAutoAssignments(sheet, workbook, registerSheetState);

        int startHour = RegisterGridLayout.clampHour(Math.min(requestedStartHour, requestedEndHour));
        int endHour = RegisterGridLayout.clampHour(Math.max(requestedStartHour, requestedEndHour));
        if (targetRegisterCount <= 0) {
            registerSheetState.getBlockedForNextRun().clear();
            return;
        }

        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        int leadershipTeamRowIndex = findLeadershipTeamRowIndex(sheet, formatter, evaluator);
        List<Integer> crewRows = collectCrewRows(sheet, leadershipTeamRowIndex, formatter, evaluator);
        if (crewRows.isEmpty()) {
            registerSheetState.getBlockedForNextRun().clear();
            return;
        }

        AutoGreedyContext greedyContext = buildAutoGreedyContext(
                sheet,
                registerSheetState,
                crewRows,
                startHour,
                endHour,
                targetRegisterCount,
                floorPrioritiesByName,
                formatter,
                evaluator
        );
        buildAutoAssignmentPlan(greedyContext);
        applyAutomaticGreedyPlan(sheet, workbook, registerSheetState, greedyContext);
        registerSheetState.getBlockedForNextRun().clear();
    }

    public boolean hasYellowAssignments(Sheet sheet) {
        if (sheet == null) {
            return false;
        }

        for (Row row : sheet) {
            short lastCellNumber = row.getLastCellNum();
            if (lastCellNumber <= 0) {
                continue;
            }

            for (int columnIndex = 0; columnIndex < lastCellNumber; columnIndex++) {
                Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                if (cell == null || !isYellowAssignmentStyle(cell.getCellStyle())) {
                    continue;
                }
                return true;
            }
        }

        return false;
    }

    public boolean clearYellowAssignments(
            Workbook workbook,
            Sheet sheet,
            RegisterSheetState registerSheetState
    ) {
        if (workbook == null || sheet == null || registerSheetState == null) {
            return false;
        }

        boolean clearedAnyAssignments = false;
        int lastRowIndex = Math.max(sheet.getLastRowNum(), RegisterGridLayout.HOUR_GRID_START_ROW_INDEX);
        for (int rowIndex = RegisterGridLayout.HOUR_GRID_START_ROW_INDEX; rowIndex <= lastRowIndex; rowIndex++) {
            for (int columnIndex = RegisterGridLayout.HOUR_GRID_START_COLUMN_INDEX;
                 columnIndex < RegisterGridLayout.HOUR_GRID_END_COLUMN_EXCLUSIVE;
                 columnIndex++) {
                Cell cell = getCell(sheet, rowIndex, columnIndex);
                if (cell == null || !isYellowAssignmentStyle(cell.getCellStyle())) {
                    continue;
                }

                restoreCellStyle(sheet, workbook, registerSheetState, rowIndex, columnIndex);
                clearedAnyAssignments = true;
            }
        }

        registerSheetState.getManualAssignments().clear();
        registerSheetState.getAutoAssignments().clear();
        registerSheetState.getBlockedForNextRun().clear();
        removeStoredRegisterAssignmentsFromCrewRowSnapshots(workbook, registerSheetState);
        return clearedAnyAssignments;
    }

    public boolean toggleCrewRowXFill(
            Workbook workbook,
            Sheet sheet,
            RegisterSheetState registerSheetState,
            int rowIndex
    ) {
        if (workbook == null || sheet == null || registerSheetState == null) {
            return false;
        }

        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        int leadershipTeamRowIndex = findLeadershipTeamRowIndex(sheet, formatter, evaluator);
        if (!isCrewRow(sheet, rowIndex, leadershipTeamRowIndex, formatter, evaluator)) {
            return false;
        }

        CrewRowXFillSnapshot crewRowXFillSnapshot = registerSheetState.getCrewRowXFillSnapshot(rowIndex);
        if (crewRowXFillSnapshot != null) {
            restoreCrewRowFromXFillSnapshot(sheet, registerSheetState, rowIndex, crewRowXFillSnapshot);
            return true;
        }

        if (isCrewRowFullyXFilled(sheet, rowIndex, formatter, evaluator)) {
            return clearCrewRowXValues(sheet, rowIndex, formatter, evaluator);
        }

        CrewRowXFillSnapshot xFillSnapshot = captureCrewRowXFillSnapshot(sheet, registerSheetState, rowIndex);
        registerSheetState.rememberCrewRowXFillSnapshot(rowIndex, xFillSnapshot);
        registerSheetState.getManualAssignments().removeIf(registerBlockKey -> registerBlockKey.rowIndex() == rowIndex);
        registerSheetState.getAutoAssignments().removeIf(registerBlockKey -> registerBlockKey.rowIndex() == rowIndex);

        boolean changed = false;
        Set<Long> processedMergedAnchors = new HashSet<>();
        for (int columnIndex = RegisterGridLayout.HOUR_GRID_START_COLUMN_INDEX;
             columnIndex < RegisterGridLayout.HOUR_GRID_END_COLUMN_EXCLUSIVE;
             columnIndex++) {
            CellRangeAddress mergedRegion = findMergedRegionContaining(sheet, rowIndex, columnIndex);
            if (mergedRegion != null) {
                if (!isSingleRowHourGridMergedRegion(mergedRegion, rowIndex)) {
                    continue;
                }

                long mergedAnchorKey = cellKey(mergedRegion.getFirstRow(), mergedRegion.getFirstColumn());
                if (!processedMergedAnchors.add(mergedAnchorKey)) {
                    continue;
                }

                Cell mergedAnchorCell = getCell(sheet, mergedRegion.getFirstRow(), mergedRegion.getFirstColumn());
                if (!isEligibleForXToggle(mergedAnchorCell)) {
                    continue;
                }

                clearCellFill(sheet, workbook, registerSheetState, mergedRegion.getFirstRow(), mergedRegion.getFirstColumn());
                if (setCellTextIfChanged(sheet, mergedRegion.getFirstRow(), mergedRegion.getFirstColumn(), "X", formatter, evaluator)) {
                    changed = true;
                }
                continue;
            }

            Cell cell = getCell(sheet, rowIndex, columnIndex);
            if (!isEligibleForXToggle(cell)) {
                continue;
            }

            clearCellFill(sheet, workbook, registerSheetState, rowIndex, columnIndex);
            if (setCellTextIfChanged(sheet, rowIndex, columnIndex, "X", formatter, evaluator)) {
                changed = true;
            }
        }

        return changed;
    }

    private AutoGreedyContext buildAutoGreedyContext(
            Sheet sheet,
            RegisterSheetState registerSheetState,
            List<Integer> crewRows,
            int startHour,
            int endHour,
            int targetRegisterCount,
            Map<String, Integer> floorPrioritiesByName,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        int crewCount = crewRows.size();
        int hourCount = endHour - startHour + 1;
        MergedRegionLookup mergedRegionLookup = buildMergedRegionLookup(sheet, crewRows);

        int[] firstColumnIndexes = new int[hourCount];
        boolean[][] baseEligible = new boolean[crewCount][hourCount];
        boolean[][] manualAssigned = new boolean[crewCount][hourCount];
        boolean[][] autoAssigned = new boolean[crewCount][hourCount];
        int[] assignedCountsByHour = new int[hourCount];
        int[] registerCountsByCrew = new int[crewCount];
        int[] floorPrioritiesByCrew = new int[crewCount];
        String[] crewMemberNames = new String[crewCount];
        boolean useFloorPriorityBias = false;

        Map<Integer, Integer> crewIndexByRow = new HashMap<>();
        for (int crewIndex = 0; crewIndex < crewRows.size(); crewIndex++) {
            int rowIndex = crewRows.get(crewIndex);
            crewIndexByRow.put(rowIndex, crewIndex);

            String crewMemberName = readCrewMemberName(sheet, rowIndex, formatter, evaluator);
            crewMemberNames[crewIndex] = crewMemberName;

            int floorPriority = normalizeFloorPriority(
                    floorPrioritiesByName.getOrDefault(normalizeCrewMemberName(crewMemberName), 0)
            );
            floorPrioritiesByCrew[crewIndex] = floorPriority;
            if (floorPriority > 0) {
                useFloorPriorityBias = true;
            }
        }

        for (RegisterBlockKey registerBlockKey : registerSheetState.getManualAssignments()) {
            Integer crewIndex = crewIndexByRow.get(registerBlockKey.rowIndex());
            if (crewIndex != null) {
                registerCountsByCrew[crewIndex]++;
            }
        }

        for (int hourOffset = 0; hourOffset < hourCount; hourOffset++) {
            int hour = startHour + hourOffset;
            int firstColumnIndex = RegisterGridLayout.hourToFirstColumn(hour);
            firstColumnIndexes[hourOffset] = firstColumnIndex;

            for (int crewIndex = 0; crewIndex < crewCount; crewIndex++) {
                int rowIndex = crewRows.get(crewIndex);
                RegisterBlockKey registerBlockKey = new RegisterBlockKey(rowIndex, firstColumnIndex);
                if (registerSheetState.getManualAssignments().contains(registerBlockKey)) {
                    manualAssigned[crewIndex][hourOffset] = true;
                    assignedCountsByHour[hourOffset]++;
                    continue;
                }

                if (registerSheetState.getBlockedForNextRun().contains(registerBlockKey)) {
                    continue;
                }

                if (registerSheetState.getManualAssignments().contains(new RegisterBlockKey(rowIndex, firstColumnIndex - 2))
                        || registerSheetState.getManualAssignments().contains(new RegisterBlockKey(rowIndex, firstColumnIndex + 2))) {
                    continue;
                }

                if (isHourBlockAvailableForAutomaticAssignment(
                        sheet,
                        mergedRegionLookup,
                        rowIndex,
                        firstColumnIndex,
                        formatter,
                        evaluator
                )) {
                    baseEligible[crewIndex][hourOffset] = true;
                }
            }
        }

        return new AutoGreedyContext(
                crewRows,
                firstColumnIndexes,
                baseEligible,
                manualAssigned,
                autoAssigned,
                assignedCountsByHour,
                registerCountsByCrew,
                targetRegisterCount,
                crewMemberNames,
                floorPrioritiesByCrew,
                useFloorPriorityBias,
                useFloorPriorityBias ? 1 : 0
        );
    }

    private void buildAutoAssignmentPlan(AutoGreedyContext greedyContext) {
        boolean progressMade;
        do {
            progressMade = false;

            int selectedHourOffset = selectMostConstrainedHour(greedyContext);
            if (selectedHourOffset < 0) {
                break;
            }

            int directCrewIndex = selectBestDirectCrew(greedyContext, selectedHourOffset);
            if (directCrewIndex >= 0) {
                assignAuto(greedyContext, directCrewIndex, selectedHourOffset);
                progressMade = true;
                continue;
            }

            if (tryRepairHour(greedyContext, selectedHourOffset, new boolean[greedyContext.hourCount()], 0, 5)) {
                progressMade = true;
            }
        } while (progressMade);

        boolean repairedAnyHour;
        do {
            repairedAnyHour = false;

            for (int hourOffset = 0; hourOffset < greedyContext.hourCount(); hourOffset++) {
                while (needsMoreAssignments(greedyContext, hourOffset)
                        && tryRepairHour(greedyContext, hourOffset, new boolean[greedyContext.hourCount()], 0, 6)) {
                    repairedAnyHour = true;
                }
            }
        } while (repairedAnyHour);
    }

    private int selectMostConstrainedHour(AutoGreedyContext greedyContext) {
        int selectedHourOffset = -1;
        int lowestCandidateCount = Integer.MAX_VALUE;
        int highestRemainingCapacity = -1;

        for (int hourOffset = 0; hourOffset < greedyContext.hourCount(); hourOffset++) {
            if (!needsMoreAssignments(greedyContext, hourOffset)) {
                continue;
            }

            int candidateCount = countDirectCandidates(greedyContext, hourOffset);
            int remainingCapacity = greedyContext.targetRegisterCount() - greedyContext.assignedCountsByHour()[hourOffset];
            if (candidateCount < lowestCandidateCount
                    || (candidateCount == lowestCandidateCount && remainingCapacity > highestRemainingCapacity)
                    || (candidateCount == lowestCandidateCount
                    && remainingCapacity == highestRemainingCapacity
                    && hourOffset < selectedHourOffset)) {
                selectedHourOffset = hourOffset;
                lowestCandidateCount = candidateCount;
                highestRemainingCapacity = remainingCapacity;
            }
        }

        return selectedHourOffset;
    }

    private int countDirectCandidates(AutoGreedyContext greedyContext, int hourOffset) {
        int candidateCount = 0;
        for (int crewIndex = 0; crewIndex < greedyContext.crewRows().size(); crewIndex++) {
            if (isDirectlyAssignable(greedyContext, crewIndex, hourOffset)) {
                candidateCount++;
            }
        }
        return candidateCount;
    }

    private int selectBestDirectCrew(AutoGreedyContext greedyContext, int hourOffset) {
        int bestCrewIndex = -1;

        for (int crewIndex = 0; crewIndex < greedyContext.crewRows().size(); crewIndex++) {
            if (!isDirectlyAssignable(greedyContext, crewIndex, hourOffset)) {
                continue;
            }

            if (bestCrewIndex < 0 || compareCrewPreference(greedyContext, crewIndex, bestCrewIndex) < 0) {
                bestCrewIndex = crewIndex;
            }
        }

        return bestCrewIndex;
    }

    private int countDirectOptionsForCrew(AutoGreedyContext greedyContext, int crewIndex) {
        int optionCount = 0;
        for (int hourOffset = 0; hourOffset < greedyContext.hourCount(); hourOffset++) {
            if (needsMoreAssignments(greedyContext, hourOffset) && isDirectlyAssignable(greedyContext, crewIndex, hourOffset)) {
                optionCount++;
            }
        }
        return optionCount;
    }

    private boolean isDirectlyAssignable(AutoGreedyContext greedyContext, int crewIndex, int hourOffset) {
        return greedyContext.baseEligible()[crewIndex][hourOffset]
                && !greedyContext.manualAssigned()[crewIndex][hourOffset]
                && !greedyContext.autoAssigned()[crewIndex][hourOffset]
                && !hasAdjacentAutoAssignment(greedyContext, crewIndex, hourOffset);
    }

    private boolean hasAdjacentAutoAssignment(AutoGreedyContext greedyContext, int crewIndex, int hourOffset) {
        return (hourOffset > 0 && greedyContext.autoAssigned()[crewIndex][hourOffset - 1])
                || (hourOffset + 1 < greedyContext.hourCount() && greedyContext.autoAssigned()[crewIndex][hourOffset + 1]);
    }

    private boolean needsMoreAssignments(AutoGreedyContext greedyContext, int hourOffset) {
        return greedyContext.assignedCountsByHour()[hourOffset] < greedyContext.targetRegisterCount();
    }

    private boolean tryRepairHour(
            AutoGreedyContext greedyContext,
            int hourOffset,
            boolean[] visitedHours,
            int depth,
            int maxDepth
    ) {
        if (!needsMoreAssignments(greedyContext, hourOffset)) {
            return true;
        }
        if (depth > maxDepth || visitedHours[hourOffset]) {
            return false;
        }

        visitedHours[hourOffset] = true;
        List<Integer> orderedCrewIndices = buildRepairCrewOrder(greedyContext, hourOffset);
        for (int crewIndex : orderedCrewIndices) {
            AutoGreedySnapshot greedySnapshot = snapshotGreedyState(greedyContext);
            List<Integer> conflictingHours = findConflictingAutoHours(greedyContext, crewIndex, hourOffset);
            if (conflictingHours == null) {
                continue;
            }

            if (conflictingHours.isEmpty()) {
                assignAuto(greedyContext, crewIndex, hourOffset);
                visitedHours[hourOffset] = false;
                return true;
            }

            for (int conflictingHourOffset : conflictingHours) {
                unassignAuto(greedyContext, crewIndex, conflictingHourOffset);
            }
            assignAuto(greedyContext, crewIndex, hourOffset);

            boolean repaired = true;
            for (int conflictingHourOffset : conflictingHours) {
                if (needsMoreAssignments(greedyContext, conflictingHourOffset)
                        && !tryRepairHour(greedyContext, conflictingHourOffset, visitedHours, depth + 1, maxDepth)) {
                    repaired = false;
                    break;
                }
            }

            if (repaired) {
                visitedHours[hourOffset] = false;
                return true;
            }

            restoreGreedyState(greedyContext, greedySnapshot);
        }

        visitedHours[hourOffset] = false;
        return false;
    }

    private List<Integer> buildRepairCrewOrder(AutoGreedyContext greedyContext, int hourOffset) {
        List<Integer> orderedCrewIndices = new ArrayList<>();
        for (int crewIndex = 0; crewIndex < greedyContext.crewRows().size(); crewIndex++) {
            if (greedyContext.baseEligible()[crewIndex][hourOffset]
                    && !greedyContext.manualAssigned()[crewIndex][hourOffset]
                    && !greedyContext.autoAssigned()[crewIndex][hourOffset]) {
                orderedCrewIndices.add(crewIndex);
            }
        }

        orderedCrewIndices.sort((leftCrewIndex, rightCrewIndex) -> {
            int leftConflicts = countConflictingAutoHours(greedyContext, leftCrewIndex, hourOffset);
            int rightConflicts = countConflictingAutoHours(greedyContext, rightCrewIndex, hourOffset);
            if (leftConflicts != rightConflicts) {
                return Integer.compare(leftConflicts, rightConflicts);
            }

            return compareCrewPreference(greedyContext, leftCrewIndex, rightCrewIndex);
        });
        return orderedCrewIndices;
    }

    private int compareCrewPreference(AutoGreedyContext greedyContext, int leftCrewIndex, int rightCrewIndex) {
        if (!greedyContext.useFloorPriorityBias()) {
            return compareDefaultCrewPreference(greedyContext, leftCrewIndex, rightCrewIndex);
        }

        boolean leftNeedsMinimum = needsMinimumAssignment(greedyContext, leftCrewIndex);
        boolean rightNeedsMinimum = needsMinimumAssignment(greedyContext, rightCrewIndex);
        if (leftNeedsMinimum != rightNeedsMinimum) {
            return leftNeedsMinimum ? -1 : 1;
        }

        if (leftNeedsMinimum) {
            int flexibilityComparison = Integer.compare(
                    countDirectOptionsForCrew(greedyContext, leftCrewIndex),
                    countDirectOptionsForCrew(greedyContext, rightCrewIndex)
            );
            if (flexibilityComparison != 0) {
                return flexibilityComparison;
            }

            int floorPriorityComparison = Integer.compare(
                    greedyContext.floorPrioritiesByCrew()[leftCrewIndex],
                    greedyContext.floorPrioritiesByCrew()[rightCrewIndex]
            );
            if (floorPriorityComparison != 0) {
                return floorPriorityComparison;
            }

            return Integer.compare(leftCrewIndex, rightCrewIndex);
        }

        int floorPriorityComparison = Integer.compare(
                greedyContext.floorPrioritiesByCrew()[leftCrewIndex],
                greedyContext.floorPrioritiesByCrew()[rightCrewIndex]
        );
        if (floorPriorityComparison != 0) {
            return floorPriorityComparison;
        }

        return compareDefaultCrewPreference(greedyContext, leftCrewIndex, rightCrewIndex);
    }

    private int compareDefaultCrewPreference(AutoGreedyContext greedyContext, int leftCrewIndex, int rightCrewIndex) {
        int leftCount = greedyContext.registerCountsByCrew()[leftCrewIndex];
        int rightCount = greedyContext.registerCountsByCrew()[rightCrewIndex];
        if (leftCount != rightCount) {
            return Integer.compare(leftCount, rightCount);
        }

        int leftFlexibility = countDirectOptionsForCrew(greedyContext, leftCrewIndex);
        int rightFlexibility = countDirectOptionsForCrew(greedyContext, rightCrewIndex);
        if (leftFlexibility != rightFlexibility) {
            return Integer.compare(leftFlexibility, rightFlexibility);
        }

        return Integer.compare(leftCrewIndex, rightCrewIndex);
    }

    private boolean needsMinimumAssignment(AutoGreedyContext greedyContext, int crewIndex) {
        return greedyContext.minimumAssignmentsPerCrew() > 0
                && greedyContext.registerCountsByCrew()[crewIndex] < greedyContext.minimumAssignmentsPerCrew();
    }

    private int countConflictingAutoHours(AutoGreedyContext greedyContext, int crewIndex, int hourOffset) {
        List<Integer> conflictingHours = findConflictingAutoHours(greedyContext, crewIndex, hourOffset);
        return conflictingHours == null ? Integer.MAX_VALUE : conflictingHours.size();
    }

    private List<Integer> findConflictingAutoHours(AutoGreedyContext greedyContext, int crewIndex, int hourOffset) {
        if (!greedyContext.baseEligible()[crewIndex][hourOffset]
                || greedyContext.manualAssigned()[crewIndex][hourOffset]
                || greedyContext.autoAssigned()[crewIndex][hourOffset]) {
            return null;
        }

        List<Integer> conflictingHours = new ArrayList<>(2);
        if (hourOffset > 0 && greedyContext.autoAssigned()[crewIndex][hourOffset - 1]) {
            conflictingHours.add(hourOffset - 1);
        }
        if (hourOffset + 1 < greedyContext.hourCount() && greedyContext.autoAssigned()[crewIndex][hourOffset + 1]) {
            conflictingHours.add(hourOffset + 1);
        }
        return conflictingHours;
    }

    private void assignAuto(AutoGreedyContext greedyContext, int crewIndex, int hourOffset) {
        if (greedyContext.autoAssigned()[crewIndex][hourOffset]) {
            return;
        }

        greedyContext.autoAssigned()[crewIndex][hourOffset] = true;
        greedyContext.assignedCountsByHour()[hourOffset]++;
        greedyContext.registerCountsByCrew()[crewIndex]++;
    }

    private void unassignAuto(AutoGreedyContext greedyContext, int crewIndex, int hourOffset) {
        if (!greedyContext.autoAssigned()[crewIndex][hourOffset]) {
            return;
        }

        greedyContext.autoAssigned()[crewIndex][hourOffset] = false;
        greedyContext.assignedCountsByHour()[hourOffset]--;
        greedyContext.registerCountsByCrew()[crewIndex]--;
    }

    private AutoGreedySnapshot snapshotGreedyState(AutoGreedyContext greedyContext) {
        boolean[][] autoAssignedSnapshot = new boolean[greedyContext.autoAssigned().length][];
        for (int crewIndex = 0; crewIndex < greedyContext.autoAssigned().length; crewIndex++) {
            autoAssignedSnapshot[crewIndex] = greedyContext.autoAssigned()[crewIndex].clone();
        }

        return new AutoGreedySnapshot(
                autoAssignedSnapshot,
                greedyContext.assignedCountsByHour().clone(),
                greedyContext.registerCountsByCrew().clone()
        );
    }

    private void restoreGreedyState(AutoGreedyContext greedyContext, AutoGreedySnapshot greedySnapshot) {
        for (int crewIndex = 0; crewIndex < greedyContext.autoAssigned().length; crewIndex++) {
            System.arraycopy(
                    greedySnapshot.autoAssigned()[crewIndex],
                    0,
                    greedyContext.autoAssigned()[crewIndex],
                    0,
                    greedyContext.autoAssigned()[crewIndex].length
            );
        }

        System.arraycopy(
                greedySnapshot.assignedCountsByHour(),
                0,
                greedyContext.assignedCountsByHour(),
                0,
                greedyContext.assignedCountsByHour().length
        );
        System.arraycopy(
                greedySnapshot.registerCountsByCrew(),
                0,
                greedyContext.registerCountsByCrew(),
                0,
                greedyContext.registerCountsByCrew().length
        );
    }

    private void applyAutomaticGreedyPlan(
            Sheet sheet,
            Workbook workbook,
            RegisterSheetState registerSheetState,
            AutoGreedyContext greedyContext
    ) {
        for (int crewIndex = 0; crewIndex < greedyContext.crewRows().size(); crewIndex++) {
            int rowIndex = greedyContext.crewRows().get(crewIndex);
            for (int hourOffset = 0; hourOffset < greedyContext.hourCount(); hourOffset++) {
                if (!greedyContext.autoAssigned()[crewIndex][hourOffset]) {
                    continue;
                }

                RegisterBlockKey registerBlockKey = new RegisterBlockKey(rowIndex, greedyContext.firstColumnIndexes()[hourOffset]);
                applyRegisterBlockFill(sheet, workbook, registerSheetState, registerBlockKey);
                registerSheetState.getAutoAssignments().add(registerBlockKey);
            }
        }
    }

    public void discardWorkbook(Workbook workbook) {
        if (workbook != null) {
            yellowStyleCacheByWorkbook.remove(workbook);
        }
    }

    private void captureBaselineStyles(Workbook workbook, Sheet sheet, RegisterSheetState registerSheetState) {
        int lastRowIndex = Math.max(sheet.getLastRowNum(), RegisterGridLayout.HOUR_GRID_START_ROW_INDEX);
        for (int rowIndex = RegisterGridLayout.HOUR_GRID_START_ROW_INDEX; rowIndex <= lastRowIndex; rowIndex++) {
            for (int columnIndex = RegisterGridLayout.HOUR_GRID_START_COLUMN_INDEX;
                 columnIndex < RegisterGridLayout.HOUR_GRID_END_COLUMN_EXCLUSIVE;
                 columnIndex++) {
                registerSheetState.rememberBaselineStyle(
                        rowIndex,
                        columnIndex,
                        resolveCurrentStyle(workbook, sheet, rowIndex, columnIndex)
                );
            }
        }
    }

    private AutoPlanningContext buildAutoPlanningContext(
            Sheet sheet,
            RegisterSheetState registerSheetState,
            List<Integer> crewRows,
            int startHour,
            int endHour,
            int targetRegisterCount,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        int hourCount = endHour - startHour + 1;
        MergedRegionLookup mergedRegionLookup = buildMergedRegionLookup(sheet, crewRows);
        Map<Integer, Integer> crewIndexByRow = new HashMap<>();
        for (int crewIndex = 0; crewIndex < crewRows.size(); crewIndex++) {
            crewIndexByRow.put(crewRows.get(crewIndex), crewIndex);
        }

        int[] firstColumnIndexes = new int[hourCount];
        BigInteger[] manualMasks = new BigInteger[hourCount];
        BigInteger[] eligibleMasks = new BigInteger[hourCount];
        BigInteger[] nextHourManualMasks = new BigInteger[hourCount];
        BigInteger[] relevantPreviousMasks = new BigInteger[hourCount];
        List<BigInteger>[] subsetOptions = new List[hourCount];
        int[] capacities = new int[hourCount];

        for (int hourOffset = 0; hourOffset < hourCount; hourOffset++) {
            int hour = startHour + hourOffset;
            int firstColumnIndex = RegisterGridLayout.hourToFirstColumn(hour);
            firstColumnIndexes[hourOffset] = firstColumnIndex;

            BigInteger manualMask = buildAssignmentMaskForHour(
                    registerSheetState.getManualAssignments(),
                    crewIndexByRow,
                    firstColumnIndex
            );
            BigInteger blockedMask = buildAssignmentMaskForHour(
                    registerSheetState.getBlockedForNextRun(),
                    crewIndexByRow,
                    firstColumnIndex
            );

            manualMasks[hourOffset] = manualMask;
            capacities[hourOffset] = Math.max(0, targetRegisterCount - manualMask.bitCount());
            eligibleMasks[hourOffset] = buildEligibleCrewMask(
                    sheet,
                    mergedRegionLookup,
                    crewRows,
                    manualMask,
                    blockedMask,
                    firstColumnIndex,
                    formatter,
                    evaluator
            );
            relevantPreviousMasks[hourOffset] = eligibleMasks[hourOffset];

            int nextHour = hour + 1;
            nextHourManualMasks[hourOffset] = nextHour > RegisterGridLayout.LAST_REGISTER_HOUR
                    ? BigInteger.ZERO
                    : buildAssignmentMaskForHour(
                    registerSheetState.getManualAssignments(),
                    crewIndexByRow,
                    RegisterGridLayout.hourToFirstColumn(nextHour)
            );
            subsetOptions[hourOffset] = buildSubsetOptions(
                    eligibleMasks[hourOffset].andNot(nextHourManualMasks[hourOffset]),
                    capacities[hourOffset]
            );
        }

        int[] remainingCapacityUpperBounds = new int[hourCount + 1];
        for (int hourOffset = hourCount - 1; hourOffset >= 0; hourOffset--) {
            remainingCapacityUpperBounds[hourOffset] = capacities[hourOffset] + remainingCapacityUpperBounds[hourOffset + 1];
        }

        BigInteger initialPreviousHourAssignments = startHour <= RegisterGridLayout.FIRST_REGISTER_HOUR
                ? BigInteger.ZERO
                : buildAssignmentMaskForHour(
                registerSheetState.getManualAssignments(),
                crewIndexByRow,
                RegisterGridLayout.hourToFirstColumn(startHour - 1)
        );

        return new AutoPlanningContext(
                crewRows,
                firstColumnIndexes,
                manualMasks,
                relevantPreviousMasks,
                subsetOptions,
                capacities,
                buildRegisterCountsByCrew(registerSheetState, crewIndexByRow),
                remainingCapacityUpperBounds,
                normalizePreviousAssignments(eligibleMasks[0], initialPreviousHourAssignments)
        );
    }

    private int[] buildRegisterCountsByCrew(
            RegisterSheetState registerSheetState,
            Map<Integer, Integer> crewIndexByRow
    ) {
        int[] registerCountsByCrew = new int[crewIndexByRow.size()];
        addRegisterCounts(registerCountsByCrew, registerSheetState.getManualAssignments(), crewIndexByRow);
        addRegisterCounts(registerCountsByCrew, registerSheetState.getAutoAssignments(), crewIndexByRow);
        return registerCountsByCrew;
    }

    private void addRegisterCounts(
            int[] registerCountsByCrew,
            Set<RegisterBlockKey> assignments,
            Map<Integer, Integer> crewIndexByRow
    ) {
        for (RegisterBlockKey registerBlockKey : assignments) {
            Integer crewIndex = crewIndexByRow.get(registerBlockKey.rowIndex());
            if (crewIndex != null) {
                registerCountsByCrew[crewIndex]++;
            }
        }
    }

    private BigInteger buildAssignmentMaskForHour(
            Set<RegisterBlockKey> assignments,
            Map<Integer, Integer> crewIndexByRow,
            int firstColumnIndex
    ) {
        BigInteger assignmentMask = BigInteger.ZERO;
        for (RegisterBlockKey registerBlockKey : assignments) {
            if (registerBlockKey.firstColumnIndex() != firstColumnIndex) {
                continue;
            }

            Integer crewIndex = crewIndexByRow.get(registerBlockKey.rowIndex());
            if (crewIndex != null) {
                assignmentMask = assignmentMask.setBit(crewIndex);
            }
        }
        return assignmentMask;
    }

    private BigInteger buildEligibleCrewMask(
            Sheet sheet,
            MergedRegionLookup mergedRegionLookup,
            List<Integer> crewRows,
            BigInteger manualMask,
            BigInteger blockedMask,
            int firstColumnIndex,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        BigInteger eligibleMask = BigInteger.ZERO;
        for (int crewIndex = 0; crewIndex < crewRows.size(); crewIndex++) {
            if (manualMask.testBit(crewIndex) || blockedMask.testBit(crewIndex)) {
                continue;
            }

            int rowIndex = crewRows.get(crewIndex);
            if (isHourBlockAvailableForAutomaticAssignment(
                    sheet,
                    mergedRegionLookup,
                    rowIndex,
                    firstColumnIndex,
                    formatter,
                    evaluator
            )) {
                eligibleMask = eligibleMask.setBit(crewIndex);
            }
        }
        return eligibleMask;
    }

    private MergedRegionLookup buildMergedRegionLookup(Sheet sheet, List<Integer> crewRows) {
        if (crewRows.isEmpty() || sheet.getNumMergedRegions() == 0) {
            return new MergedRegionLookup(new HashMap<>());
        }

        int firstCrewRowIndex = crewRows.get(0);
        int lastCrewRowIndex = crewRows.get(crewRows.size() - 1);
        Map<Long, CellRangeAddress> regionsByCell = new HashMap<>();

        for (int regionIndex = 0; regionIndex < sheet.getNumMergedRegions(); regionIndex++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(regionIndex);

            int firstRowIndex = Math.max(firstCrewRowIndex, mergedRegion.getFirstRow());
            int lastRowIndex = Math.min(lastCrewRowIndex, mergedRegion.getLastRow());
            int firstColumnIndex = Math.max(RegisterGridLayout.HOUR_GRID_START_COLUMN_INDEX, mergedRegion.getFirstColumn());
            int lastColumnIndex = Math.min(RegisterGridLayout.HOUR_GRID_END_COLUMN_EXCLUSIVE - 1, mergedRegion.getLastColumn());
            if (firstRowIndex > lastRowIndex || firstColumnIndex > lastColumnIndex) {
                continue;
            }

            for (int rowIndex = firstRowIndex; rowIndex <= lastRowIndex; rowIndex++) {
                for (int columnIndex = firstColumnIndex; columnIndex <= lastColumnIndex; columnIndex++) {
                    regionsByCell.put(cellKey(rowIndex, columnIndex), mergedRegion);
                }
            }
        }

        return new MergedRegionLookup(regionsByCell);
    }

    private List<Integer> collectCrewRows(
            Sheet sheet,
            int leadershipTeamRowIndex,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        List<Integer> crewRows = new ArrayList<>();
        int lastCrewRowIndex = Math.min(sheet.getLastRowNum(), leadershipTeamRowIndex - 1);
        for (int rowIndex = RegisterGridLayout.HOUR_GRID_START_ROW_INDEX; rowIndex <= lastCrewRowIndex; rowIndex++) {
            if (isCrewRow(sheet, rowIndex, leadershipTeamRowIndex, formatter, evaluator)) {
                crewRows.add(rowIndex);
            }
        }
        return crewRows;
    }

    private int computeMaxAssignments(
            AutoPlanningContext planningContext,
            int hourOffset,
            BigInteger previousHourAssignments,
            Map<SearchState, Integer> maxAssignmentsCache
    ) {
        if (hourOffset >= planningContext.hourCount()) {
            return 0;
        }

        BigInteger normalizedPreviousAssignments = normalizePreviousAssignments(
                planningContext.relevantPreviousMasks()[hourOffset],
                previousHourAssignments
        );
        SearchState searchState = new SearchState(hourOffset, normalizedPreviousAssignments);
        Integer cachedValue = maxAssignmentsCache.get(searchState);
        if (cachedValue != null) {
            return cachedValue;
        }

        int bestAdditionalAssignments = 0;
        for (BigInteger subsetMask : planningContext.subsetOptions()[hourOffset]) {
            if (subsetMask.and(normalizedPreviousAssignments).signum() != 0) {
                continue;
            }

            int subsetSize = subsetMask.bitCount();
            if (subsetSize + planningContext.remainingCapacityUpperBounds()[hourOffset + 1] < bestAdditionalAssignments) {
                break;
            }

            int totalAssignments = subsetSize + computeMaxAssignments(
                    planningContext,
                    hourOffset + 1,
                    normalizePreviousAssignmentsForNextHour(
                            planningContext,
                            hourOffset,
                            planningContext.manualMasks()[hourOffset].or(subsetMask)
                    ),
                    maxAssignmentsCache
            );
            if (totalAssignments > bestAdditionalAssignments) {
                bestAdditionalAssignments = totalAssignments;
            }
        }

        maxAssignmentsCache.put(searchState, bestAdditionalAssignments);
        return bestAdditionalAssignments;
    }

    private BigInteger[] buildOptimalAutoMasks(
            AutoPlanningContext planningContext,
            Map<SearchState, Integer> maxAssignmentsCache
    ) {
        BigInteger[] chosenAutoMasks = new BigInteger[planningContext.hourCount()];
        int[] registerCountsByCrew = planningContext.startingRegisterCounts().clone();
        BigInteger previousHourAssignments = planningContext.initialPreviousHourAssignments();

        for (int hourOffset = 0; hourOffset < planningContext.hourCount(); hourOffset++) {
            int optimalAssignmentsFromState = computeMaxAssignments(
                    planningContext,
                    hourOffset,
                    previousHourAssignments,
                    maxAssignmentsCache
            );

            BigInteger normalizedPreviousAssignments = normalizePreviousAssignments(
                    planningContext.relevantPreviousMasks()[hourOffset],
                    previousHourAssignments
            );

            BigInteger bestSubsetMask = BigInteger.ZERO;
            int bestSubsetFairnessCost = Integer.MAX_VALUE;
            int bestSubsetSize = -1;

            for (BigInteger subsetMask : planningContext.subsetOptions()[hourOffset]) {
                if (subsetMask.and(normalizedPreviousAssignments).signum() != 0) {
                    continue;
                }

                int subsetSize = subsetMask.bitCount();
                if (subsetSize + planningContext.remainingCapacityUpperBounds()[hourOffset + 1] < optimalAssignmentsFromState) {
                    break;
                }

                int totalAssignments = subsetSize + computeMaxAssignments(
                        planningContext,
                        hourOffset + 1,
                        normalizePreviousAssignmentsForNextHour(
                                planningContext,
                                hourOffset,
                                planningContext.manualMasks()[hourOffset].or(subsetMask)
                        ),
                        maxAssignmentsCache
                );
                if (totalAssignments != optimalAssignmentsFromState) {
                    continue;
                }

                int fairnessCost = incrementalFairnessCost(registerCountsByCrew, subsetMask);
                if (subsetSize > bestSubsetSize
                        || (subsetSize == bestSubsetSize && fairnessCost < bestSubsetFairnessCost)
                        || (subsetSize == bestSubsetSize
                        && fairnessCost == bestSubsetFairnessCost
                        && compareSubsetMasks(subsetMask, bestSubsetMask) < 0)) {
                    bestSubsetMask = subsetMask;
                    bestSubsetFairnessCost = fairnessCost;
                    bestSubsetSize = subsetSize;
                }
            }

            chosenAutoMasks[hourOffset] = bestSubsetMask;
            for (int crewIndex : extractSetBitIndices(bestSubsetMask)) {
                registerCountsByCrew[crewIndex]++;
            }
            previousHourAssignments = normalizePreviousAssignmentsForNextHour(
                    planningContext,
                    hourOffset,
                    planningContext.manualMasks()[hourOffset].or(bestSubsetMask)
            );
        }

        return chosenAutoMasks;
    }

    private BigInteger normalizePreviousAssignments(BigInteger relevantPreviousMask, BigInteger previousHourAssignments) {
        return previousHourAssignments.and(relevantPreviousMask);
    }

    private BigInteger normalizePreviousAssignmentsForNextHour(
            AutoPlanningContext planningContext,
            int currentHourOffset,
            BigInteger previousHourAssignments
    ) {
        int nextHourOffset = currentHourOffset + 1;
        if (nextHourOffset >= planningContext.hourCount()) {
            return BigInteger.ZERO;
        }
        return normalizePreviousAssignments(
                planningContext.relevantPreviousMasks()[nextHourOffset],
                previousHourAssignments
        );
    }

    private void applyAutomaticPlan(
            Sheet sheet,
            Workbook workbook,
            RegisterSheetState registerSheetState,
            AutoPlanningContext planningContext,
            BigInteger[] chosenAutoMasks
    ) {
        for (int hourOffset = 0; hourOffset < chosenAutoMasks.length; hourOffset++) {
            BigInteger chosenMask = chosenAutoMasks[hourOffset];
            int firstColumnIndex = planningContext.firstColumnIndexes()[hourOffset];
            for (int crewIndex : extractSetBitIndices(chosenMask)) {
                int rowIndex = planningContext.crewRows().get(crewIndex);
                RegisterBlockKey registerBlockKey = new RegisterBlockKey(rowIndex, firstColumnIndex);
                applyRegisterBlockFill(sheet, workbook, registerSheetState, registerBlockKey);
                registerSheetState.getAutoAssignments().add(registerBlockKey);
            }
        }
    }

    private int incrementalFairnessCost(int[] registerCountsByCrew, BigInteger subsetMask) {
        int fairnessCost = 0;
        for (int crewIndex : extractSetBitIndices(subsetMask)) {
            fairnessCost += (registerCountsByCrew[crewIndex] * 2) + 1;
        }
        return fairnessCost;
    }

    private int compareSubsetMasks(BigInteger leftMask, BigInteger rightMask) {
        if (leftMask.equals(rightMask)) {
            return 0;
        }

        List<Integer> leftBits = extractSetBitIndices(leftMask);
        List<Integer> rightBits = extractSetBitIndices(rightMask);
        int sharedLength = Math.min(leftBits.size(), rightBits.size());
        for (int index = 0; index < sharedLength; index++) {
            int comparison = Integer.compare(leftBits.get(index), rightBits.get(index));
            if (comparison != 0) {
                return comparison;
            }
        }
        return Integer.compare(leftBits.size(), rightBits.size());
    }

    private List<BigInteger> buildSubsetOptions(BigInteger availableMask, int capacity) {
        List<BigInteger> subsetOptions = new ArrayList<>();
        int maxSubsetSize = Math.min(capacity, availableMask.bitCount());
        List<Integer> candidateIndices = extractSetBitIndices(availableMask);

        subsetOptions.add(BigInteger.ZERO);
        for (int subsetSize = maxSubsetSize; subsetSize >= 1; subsetSize--) {
            visitExactSubsets(candidateIndices, 0, subsetSize, BigInteger.ZERO, subsetOptions::add);
        }

        subsetOptions.sort((leftMask, rightMask) -> {
            int leftSize = leftMask.bitCount();
            int rightSize = rightMask.bitCount();
            if (leftSize != rightSize) {
                return Integer.compare(rightSize, leftSize);
            }
            return compareSubsetMasks(leftMask, rightMask);
        });
        return subsetOptions;
    }

    private List<Integer> extractSetBitIndices(BigInteger mask) {
        List<Integer> setBitIndices = new ArrayList<>();
        BigInteger remainingMask = mask;
        while (remainingMask.signum() > 0) {
            int lowestSetBit = remainingMask.getLowestSetBit();
            setBitIndices.add(lowestSetBit);
            remainingMask = remainingMask.clearBit(lowestSetBit);
        }
        return setBitIndices;
    }

    private void visitExactSubsets(
            List<Integer> candidateIndices,
            int startIndex,
            int remainingToChoose,
            BigInteger currentMask,
            MaskVisitor maskVisitor
    ) {
        if (remainingToChoose == 0) {
            maskVisitor.accept(currentMask);
            return;
        }

        if (candidateIndices.size() - startIndex < remainingToChoose) {
            return;
        }

        for (int index = startIndex; index <= candidateIndices.size() - remainingToChoose; index++) {
            visitExactSubsets(
                    candidateIndices,
                    index + 1,
                    remainingToChoose - 1,
                    currentMask.setBit(candidateIndices.get(index)),
                    maskVisitor
            );
        }
    }

    private boolean isCellAvailableForAutomaticAssignment(
            Sheet sheet,
            MergedRegionLookup mergedRegionLookup,
            int rowIndex,
            int columnIndex,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        CellRangeAddress mergedRegion = findMergedRegionContaining(mergedRegionLookup, rowIndex, columnIndex);
        int resolvedRowIndex = mergedRegion == null ? rowIndex : mergedRegion.getFirstRow();
        int resolvedColumnIndex = mergedRegion == null ? columnIndex : mergedRegion.getFirstColumn();

        Cell cell = getCell(sheet, resolvedRowIndex, resolvedColumnIndex);
        if (cell == null) {
            return true;
        }

        String displayValue = formatter.formatCellValue(cell, evaluator);
        if (!displayValue.isBlank()) {
            return false;
        }

        return hasNoFill(cell.getCellStyle());
    }

    private boolean isHourBlockAvailableForAutomaticAssignment(
            Sheet sheet,
            MergedRegionLookup mergedRegionLookup,
            int rowIndex,
            int firstColumnIndex,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        if (isHourBlockBlockedByMergedRegion(mergedRegionLookup, rowIndex, firstColumnIndex)) {
            return false;
        }

        boolean firstHalfAvailable = isCellAvailableForAutomaticAssignment(
                sheet,
                mergedRegionLookup,
                rowIndex,
                firstColumnIndex,
                formatter,
                evaluator
        );
        boolean secondHalfAvailable = isCellAvailableForAutomaticAssignment(
                sheet,
                mergedRegionLookup,
                rowIndex,
                firstColumnIndex + 1,
                formatter,
                evaluator
        );

        if (firstHalfAvailable != secondHalfAvailable) {
            return false;
        }

        return firstHalfAvailable;
    }

    private boolean isHourBlockBlockedByMergedRegion(
            MergedRegionLookup mergedRegionLookup,
            int rowIndex,
            int firstColumnIndex
    ) {
        CellRangeAddress firstMergedRegion = findMergedRegionContaining(mergedRegionLookup, rowIndex, firstColumnIndex);
        if (mergedRegionExtendsOutsideHourBlock(firstMergedRegion, rowIndex, firstColumnIndex)) {
            return true;
        }

        CellRangeAddress secondMergedRegion = findMergedRegionContaining(mergedRegionLookup, rowIndex, firstColumnIndex + 1);
        return mergedRegionExtendsOutsideHourBlock(secondMergedRegion, rowIndex, firstColumnIndex);
    }

    private boolean mergedRegionExtendsOutsideHourBlock(
            CellRangeAddress mergedRegion,
            int rowIndex,
            int firstColumnIndex
    ) {
        return mergedRegion != null
                && (mergedRegion.getFirstRow() != rowIndex
                || mergedRegion.getLastRow() != rowIndex
                || mergedRegion.getFirstColumn() < firstColumnIndex
                || mergedRegion.getLastColumn() > firstColumnIndex + 1);
    }

    private CellRangeAddress findMergedRegionContaining(
            MergedRegionLookup mergedRegionLookup,
            int rowIndex,
            int columnIndex
    ) {
        return mergedRegionLookup.regionsByCell().get(cellKey(rowIndex, columnIndex));
    }

    private CellRangeAddress findMergedRegionContaining(
            Sheet sheet,
            int rowIndex,
            int columnIndex
    ) {
        for (int mergedRegionIndex = 0; mergedRegionIndex < sheet.getNumMergedRegions(); mergedRegionIndex++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(mergedRegionIndex);
            if (mergedRegion.isInRange(rowIndex, columnIndex)) {
                return mergedRegion;
            }
        }
        return null;
    }

    private boolean isSingleRowHourGridMergedRegion(CellRangeAddress mergedRegion, int rowIndex) {
        return mergedRegion.getFirstRow() == rowIndex
                && mergedRegion.getLastRow() == rowIndex
                && mergedRegion.getFirstColumn() >= RegisterGridLayout.HOUR_GRID_START_COLUMN_INDEX
                && mergedRegion.getLastColumn() < RegisterGridLayout.HOUR_GRID_END_COLUMN_EXCLUSIVE;
    }

    private String readCrewMemberName(
            Sheet sheet,
            int rowIndex,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }

        Cell nameCell = row.getCell(0, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
        if (nameCell == null) {
            return "";
        }

        return formatter.formatCellValue(nameCell, evaluator).trim();
    }

    private boolean isCrewRow(
            Sheet sheet,
            int rowIndex,
            int leadershipTeamRowIndex,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        if (rowIndex >= leadershipTeamRowIndex) {
            return false;
        }

        Row row = sheet.getRow(rowIndex);
        if (row == null || row.getZeroHeight()) {
            return false;
        }

        for (int columnIndex = 0; columnIndex < RegisterGridLayout.HOUR_GRID_START_COLUMN_INDEX; columnIndex++) {
            Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
            if (cell != null && !formatter.formatCellValue(cell, evaluator).isBlank()) {
                return true;
            }
        }

        return false;
    }

    private int findLeadershipTeamRowIndex(
            Sheet sheet,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        for (int rowIndex = RegisterGridLayout.HOUR_GRID_START_ROW_INDEX; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }

            short lastCellNumber = row.getLastCellNum();
            if (lastCellNumber < 0) {
                continue;
            }

            for (int columnIndex = 0; columnIndex < lastCellNumber; columnIndex++) {
                Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                if (cell == null) {
                    continue;
                }

                String displayValue = formatter.formatCellValue(cell, evaluator).trim();
                if (LEADERSHIP_TEAM_MARKER.equals(displayValue)) {
                    return rowIndex;
                }
            }
        }

        return sheet.getLastRowNum() + 1;
    }

    private String normalizeCrewMemberName(String crewMemberName) {
        return crewMemberName == null ? "" : crewMemberName.trim().toLowerCase(Locale.ROOT);
    }

    private int normalizeFloorPriority(int floorPriority) {
        return Math.max(0, Math.min(3, floorPriority));
    }

    private void clearAutoAssignments(
            Sheet sheet,
            Workbook workbook,
            RegisterSheetState registerSheetState
    ) {
        List<RegisterBlockKey> autoAssignments = new ArrayList<>(registerSheetState.getAutoAssignments());
        for (RegisterBlockKey registerBlockKey : autoAssignments) {
            restoreRegisterBlock(sheet, workbook, registerSheetState, registerBlockKey);
        }
        registerSheetState.getAutoAssignments().clear();
    }

    private void applyRegisterBlockFill(
            Sheet sheet,
            Workbook workbook,
            RegisterSheetState registerSheetState,
            RegisterBlockKey registerBlockKey
    ) {
        applyRegisterFillToCell(sheet, workbook, registerSheetState, registerBlockKey.rowIndex(), registerBlockKey.firstColumnIndex());
        applyRegisterFillToCell(sheet, workbook, registerSheetState, registerBlockKey.rowIndex(), registerBlockKey.secondColumnIndex());
    }

    private void restoreRegisterBlock(
            Sheet sheet,
            Workbook workbook,
            RegisterSheetState registerSheetState,
            RegisterBlockKey registerBlockKey
    ) {
        restoreCellStyle(sheet, workbook, registerSheetState, registerBlockKey.rowIndex(), registerBlockKey.firstColumnIndex());
        restoreCellStyle(sheet, workbook, registerSheetState, registerBlockKey.rowIndex(), registerBlockKey.secondColumnIndex());
    }

    private void applyRegisterFillToCell(
            Sheet sheet,
            Workbook workbook,
            RegisterSheetState registerSheetState,
            int rowIndex,
            int columnIndex
    ) {
        CellStyle baselineStyle = baselineStyle(workbook, registerSheetState, rowIndex, columnIndex);
        Cell cell = getOrCreateCell(sheet, rowIndex, columnIndex);
        cell.setCellStyle(resolveYellowStyle(workbook, baselineStyle));
    }

    private void restoreCellStyle(
            Sheet sheet,
            Workbook workbook,
            RegisterSheetState registerSheetState,
            int rowIndex,
            int columnIndex
    ) {
        Cell cell = getOrCreateCell(sheet, rowIndex, columnIndex);
        cell.setCellStyle(baselineStyle(workbook, registerSheetState, rowIndex, columnIndex));
    }

    private void clearCellFill(
            Sheet sheet,
            Workbook workbook,
            RegisterSheetState registerSheetState,
            int rowIndex,
            int columnIndex
    ) {
        Cell cell = getOrCreateCell(sheet, rowIndex, columnIndex);
        CellStyle currentStyle = cell.getCellStyle();
        if (!isYellowAssignmentStyle(currentStyle)) {
            return;
        }

        cell.setCellStyle(baselineStyle(workbook, registerSheetState, rowIndex, columnIndex));
    }

    private CellStyle baselineStyle(
            Workbook workbook,
            RegisterSheetState registerSheetState,
            int rowIndex,
            int columnIndex
    ) {
        CellStyle cellStyle = registerSheetState.getBaselineStyle(rowIndex, columnIndex);
        if (cellStyle != null) {
            return cellStyle;
        }

        CellStyle resolvedStyle = resolveCurrentStyle(workbook, null, rowIndex, columnIndex);
        registerSheetState.rememberBaselineStyle(rowIndex, columnIndex, resolvedStyle);
        return resolvedStyle;
    }

    private CellStyle resolveCurrentStyle(Workbook workbook, Sheet sheet, int rowIndex, int columnIndex) {
        if (sheet == null) {
            return workbook.getCellStyleAt(0);
        }

        Cell cell = getCell(sheet, rowIndex, columnIndex);
        if (cell == null || cell.getCellStyle() == null) {
            return workbook.getCellStyleAt(0);
        }
        return cell.getCellStyle();
    }

    private CellStyle resolveYellowStyle(Workbook workbook, CellStyle baselineStyle) {
        short baselineStyleIndex = baselineStyle.getIndex();
        Map<Short, CellStyle> yellowStyleCache = yellowStyleCacheByWorkbook.computeIfAbsent(
                workbook,
                ignored -> new HashMap<>()
        );

        return yellowStyleCache.computeIfAbsent(baselineStyleIndex, ignored -> {
            CellStyle yellowStyle = workbook.createCellStyle();
            yellowStyle.cloneStyleFrom(baselineStyle);
            yellowStyle.setFillForegroundColor(REGISTER_FILL_COLOR_INDEX);
            yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            return yellowStyle;
        });
    }

    private boolean isYellowAssignmentStyle(CellStyle cellStyle) {
        return cellStyle != null
                && cellStyle.getFillPattern() == FillPatternType.SOLID_FOREGROUND
                && cellStyle.getFillForegroundColor() == REGISTER_FILL_COLOR_INDEX;
    }

    private boolean isEligibleForXToggle(Cell cell) {
        return cell == null
                || hasNoFill(cell.getCellStyle())
                || isYellowAssignmentStyle(cell.getCellStyle());
    }

    private boolean hasNoFill(CellStyle cellStyle) {
        return cellStyle == null || cellStyle.getFillPattern() == FillPatternType.NO_FILL;
    }

    private CrewRowXFillSnapshot captureCrewRowXFillSnapshot(
            Sheet sheet,
            RegisterSheetState registerSheetState,
            int rowIndex
    ) {
        Map<CellCoordinate, CellSnapshot> cellSnapshots = new LinkedHashMap<>();
        for (int columnIndex = RegisterGridLayout.HOUR_GRID_START_COLUMN_INDEX;
             columnIndex < RegisterGridLayout.HOUR_GRID_END_COLUMN_EXCLUSIVE;
             columnIndex++) {
            cellSnapshots.put(
                    new CellCoordinate(rowIndex, columnIndex),
                    captureCellSnapshot(sheet, rowIndex, columnIndex)
            );
        }

        Set<RegisterBlockKey> manualAssignments = new HashSet<>();
        for (RegisterBlockKey registerBlockKey : registerSheetState.getManualAssignments()) {
            if (registerBlockKey.rowIndex() == rowIndex) {
                manualAssignments.add(registerBlockKey);
            }
        }

        Set<RegisterBlockKey> autoAssignments = new HashSet<>();
        for (RegisterBlockKey registerBlockKey : registerSheetState.getAutoAssignments()) {
            if (registerBlockKey.rowIndex() == rowIndex) {
                autoAssignments.add(registerBlockKey);
            }
        }

        return new CrewRowXFillSnapshot(cellSnapshots, manualAssignments, autoAssignments);
    }

    private CellSnapshot captureCellSnapshot(Sheet sheet, int rowIndex, int columnIndex) {
        Cell cell = getCell(sheet, rowIndex, columnIndex);
        if (cell == null) {
            return new CellSnapshot(true, null, CellType.BLANK, null, null, null, null, null);
        }

        CellType cellType = cell.getCellType();
        return switch (cellType) {
            case STRING -> new CellSnapshot(
                    false,
                    cell.getCellStyle(),
                    cellType,
                    cell.getStringCellValue(),
                    null,
                    null,
                    null,
                    null
            );
            case NUMERIC -> new CellSnapshot(
                    false,
                    cell.getCellStyle(),
                    cellType,
                    null,
                    cell.getNumericCellValue(),
                    null,
                    null,
                    null
            );
            case BOOLEAN -> new CellSnapshot(
                    false,
                    cell.getCellStyle(),
                    cellType,
                    null,
                    null,
                    cell.getBooleanCellValue(),
                    null,
                    null
            );
            case FORMULA -> new CellSnapshot(
                    false,
                    cell.getCellStyle(),
                    cellType,
                    null,
                    null,
                    null,
                    cell.getCellFormula(),
                    null
            );
            case ERROR -> new CellSnapshot(
                    false,
                    cell.getCellStyle(),
                    cellType,
                    null,
                    null,
                    null,
                    null,
                    cell.getErrorCellValue()
            );
            case BLANK, _NONE -> new CellSnapshot(
                    false,
                    cell.getCellStyle(),
                    CellType.BLANK,
                    null,
                    null,
                    null,
                    null,
                    null
            );
        };
    }

    private void restoreCrewRowFromXFillSnapshot(
            Sheet sheet,
            RegisterSheetState registerSheetState,
            int rowIndex,
            CrewRowXFillSnapshot crewRowXFillSnapshot
    ) {
        for (Map.Entry<CellCoordinate, CellSnapshot> entry : crewRowXFillSnapshot.cellSnapshots().entrySet()) {
            restoreCellSnapshot(sheet, entry.getKey(), entry.getValue());
        }

        registerSheetState.getManualAssignments().removeIf(registerBlockKey -> registerBlockKey.rowIndex() == rowIndex);
        registerSheetState.getAutoAssignments().removeIf(registerBlockKey -> registerBlockKey.rowIndex() == rowIndex);
        registerSheetState.getManualAssignments().addAll(crewRowXFillSnapshot.manualAssignments());
        registerSheetState.getAutoAssignments().addAll(crewRowXFillSnapshot.autoAssignments());
        registerSheetState.removeCrewRowXFillSnapshot(rowIndex);
    }

    private void restoreCellSnapshot(
            Sheet sheet,
            CellCoordinate cellCoordinate,
            CellSnapshot cellSnapshot
    ) {
        if (cellSnapshot.missing()) {
            removeCell(sheet, cellCoordinate.rowIndex(), cellCoordinate.columnIndex());
            return;
        }

        Cell cell = getOrCreateCell(sheet, cellCoordinate.rowIndex(), cellCoordinate.columnIndex());
        if (cellSnapshot.cellStyle() != null) {
            cell.setCellStyle(cellSnapshot.cellStyle());
        }

        switch (cellSnapshot.cellType()) {
            case STRING -> cell.setCellValue(cellSnapshot.stringValue() == null ? "" : cellSnapshot.stringValue());
            case NUMERIC -> cell.setCellValue(cellSnapshot.numericValue() == null ? 0d : cellSnapshot.numericValue());
            case BOOLEAN -> cell.setCellValue(Boolean.TRUE.equals(cellSnapshot.booleanValue()));
            case FORMULA -> cell.setCellFormula(cellSnapshot.formula() == null ? "" : cellSnapshot.formula());
            case ERROR -> cell.setCellErrorValue(cellSnapshot.errorValue() == null ? (byte) 0 : cellSnapshot.errorValue());
            case BLANK, _NONE -> cell.setBlank();
        }
    }

    private boolean isCrewRowFullyXFilled(
            Sheet sheet,
            int rowIndex,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        boolean foundAnyTargetCell = false;
        for (int columnIndex = RegisterGridLayout.HOUR_GRID_START_COLUMN_INDEX;
             columnIndex < RegisterGridLayout.HOUR_GRID_END_COLUMN_EXCLUSIVE;
             columnIndex++) {
            Cell cell = getCell(sheet, rowIndex, columnIndex);
            if (!isEligibleForXToggle(cell)) {
                continue;
            }

            foundAnyTargetCell = true;
            if (cell == null || !hasNoFill(cell.getCellStyle())) {
                return false;
            }

            if (!"X".equals(formatter.formatCellValue(cell, evaluator))) {
                return false;
            }
        }

        return foundAnyTargetCell;
    }

    private boolean clearCrewRowXValues(
            Sheet sheet,
            int rowIndex,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        boolean changed = false;
        for (int columnIndex = RegisterGridLayout.HOUR_GRID_START_COLUMN_INDEX;
             columnIndex < RegisterGridLayout.HOUR_GRID_END_COLUMN_EXCLUSIVE;
             columnIndex++) {
            Cell cell = getCell(sheet, rowIndex, columnIndex);
            if (cell == null || !hasNoFill(cell.getCellStyle())) {
                continue;
            }

            if (!"X".equals(formatter.formatCellValue(cell, evaluator))) {
                continue;
            }

            cell.setBlank();
            changed = true;
        }

        return changed;
    }

    private void removeStoredRegisterAssignmentsFromCrewRowSnapshots(
            Workbook workbook,
            RegisterSheetState registerSheetState
    ) {
        for (Map.Entry<Integer, CrewRowXFillSnapshot> entry : new ArrayList<>(registerSheetState.getCrewRowXFillSnapshots().entrySet())) {
            int rowIndex = entry.getKey();
            CrewRowXFillSnapshot crewRowXFillSnapshot = entry.getValue();
            Map<CellCoordinate, CellSnapshot> sanitizedCellSnapshots = new LinkedHashMap<>();
            for (Map.Entry<CellCoordinate, CellSnapshot> cellEntry : crewRowXFillSnapshot.cellSnapshots().entrySet()) {
                CellCoordinate cellCoordinate = cellEntry.getKey();
                CellSnapshot cellSnapshot = cellEntry.getValue();
                CellStyle sanitizedStyle = cellSnapshot.cellStyle();
                if (isYellowAssignmentStyle(sanitizedStyle)) {
                    sanitizedStyle = baselineStyle(workbook, registerSheetState, cellCoordinate.rowIndex(), cellCoordinate.columnIndex());
                }
                sanitizedCellSnapshots.put(
                        cellCoordinate,
                        new CellSnapshot(
                                cellSnapshot.missing(),
                                sanitizedStyle,
                                cellSnapshot.cellType(),
                                cellSnapshot.stringValue(),
                                cellSnapshot.numericValue(),
                                cellSnapshot.booleanValue(),
                                cellSnapshot.formula(),
                                cellSnapshot.errorValue()
                        )
                );
            }

            registerSheetState.rememberCrewRowXFillSnapshot(
                    rowIndex,
                    new CrewRowXFillSnapshot(
                            sanitizedCellSnapshots,
                            Set.of(),
                            Set.of()
                    )
            );
        }
    }

    private boolean setCellTextIfChanged(
            Sheet sheet,
            int rowIndex,
            int columnIndex,
            String value,
            DataFormatter formatter,
            FormulaEvaluator evaluator
    ) {
        Cell existingCell = getCell(sheet, rowIndex, columnIndex);
        if (existingCell != null && value.equals(formatter.formatCellValue(existingCell, evaluator))) {
            return false;
        }

        Cell cell = getOrCreateCell(sheet, rowIndex, columnIndex);
        cell.setCellValue(value);
        return true;
    }

    private Cell getCell(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return null;
        }
        return row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
    }

    private void removeCell(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return;
        }

        Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
        if (cell == null) {
            return;
        }

        row.removeCell(cell);
    }

    private Cell getOrCreateCell(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        return cell;
    }

    private long cellKey(int rowIndex, int columnIndex) {
        return (((long) rowIndex) << 32) | (columnIndex & 0xffffffffL);
    }

    @FunctionalInterface
    private interface MaskVisitor {
        void accept(BigInteger mask);
    }

    private record SearchState(int hourOffset, BigInteger previousHourAssignments) {
    }

    private record AutoGreedyContext(
            List<Integer> crewRows,
            int[] firstColumnIndexes,
            boolean[][] baseEligible,
            boolean[][] manualAssigned,
            boolean[][] autoAssigned,
            int[] assignedCountsByHour,
            int[] registerCountsByCrew,
            int targetRegisterCount,
            String[] crewMemberNames,
            int[] floorPrioritiesByCrew,
            boolean useFloorPriorityBias,
            int minimumAssignmentsPerCrew
    ) {
        private int hourCount() {
            return firstColumnIndexes.length;
        }
    }

    private record AutoGreedySnapshot(
            boolean[][] autoAssigned,
            int[] assignedCountsByHour,
            int[] registerCountsByCrew
    ) {
    }

    private record AutoPlanningContext(
            List<Integer> crewRows,
            int[] firstColumnIndexes,
            BigInteger[] manualMasks,
            BigInteger[] relevantPreviousMasks,
            List<BigInteger>[] subsetOptions,
            int[] capacities,
            int[] startingRegisterCounts,
            int[] remainingCapacityUpperBounds,
            BigInteger initialPreviousHourAssignments
    ) {
        private int hourCount() {
            return firstColumnIndexes.length;
        }
    }

    private record MergedRegionLookup(Map<Long, CellRangeAddress> regionsByCell) {
    }
}
