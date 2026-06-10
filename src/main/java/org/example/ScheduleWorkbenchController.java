package org.example;

import javafx.animation.PauseTransition;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.PasswordField;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.Slider;
import javafx.scene.control.TablePosition;
import javafx.scene.control.TextField;
import javafx.scene.control.TextFormatter;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
import javafx.scene.input.MouseEvent;
import javafx.scene.input.MouseButton;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Window;
import javafx.util.Duration;
import javafx.util.StringConverter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.controlsfx.control.RangeSlider;
import org.controlsfx.control.spreadsheet.GridChange;
import org.controlsfx.control.spreadsheet.SpreadsheetView;
import org.controlsfx.control.spreadsheet.SpreadsheetViewSelectionModel;
import org.example.model.CrewAvailabilitySummary;
import org.example.model.DaySheetOption;
import org.example.model.LocalAccount;
import org.example.model.RegisterBlockKey;
import org.example.model.RegisterSheetState;
import org.example.service.CrewFloorPriorityService;
import org.example.service.CrewAvailabilityStatusService;
import org.example.service.CrewMemberDirectoryService;
import org.example.service.FilledWorkbookExportService;
import org.example.service.LocalAccountService;
import org.example.service.RegisterAssignmentService;
import org.example.service.RegisterGridLayout;
import org.example.service.WorkbookInspectorService;
import org.example.service.WorkbookSpreadsheetViewService;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

public class ScheduleWorkbenchController {

    private static final Duration REGISTER_TOGGLE_DELAY = Duration.millis(0);
    private static final Duration INLINE_EDIT_REVERT_WINDOW = Duration.millis(350);

    private final LocalAccountService localAccountService = new LocalAccountService();
    private final CrewMemberDirectoryService crewMemberDirectoryService = new CrewMemberDirectoryService();
    private final CrewFloorPriorityService crewFloorPriorityService = new CrewFloorPriorityService();
    private final WorkbookInspectorService workbookInspectorService = new WorkbookInspectorService();
    private final WorkbookSpreadsheetViewService workbookSpreadsheetViewService = new WorkbookSpreadsheetViewService();
    private final RegisterAssignmentService registerAssignmentService = new RegisterAssignmentService();
    private final CrewAvailabilityStatusService crewAvailabilityStatusService = new CrewAvailabilityStatusService();
    private final FilledWorkbookExportService filledWorkbookExportService = new FilledWorkbookExportService();

    private LocalAccount currentAccount;
    private Path loadedWorkbookPath;
    private Workbook loadedWorkbook;
    private final Map<Integer, RegisterSheetState> registerSheetStates = new HashMap<>();
    private PauseTransition pendingRegisterToggle;
    private PendingRegisterToggle pendingRegisterToggleRequest;
    private PendingInlineEditAdvance pendingInlineEditAdvance;
    private RecentRegisterToggle recentRegisterToggle;

    @FXML
    private Label accountStatusLabel;

    @FXML
    private HBox accountActionBox;

    @FXML
    private Button loginButton;

    @FXML
    private Button createAccountButton;

    @FXML
    private Button logoutButton;

    @FXML
    private Button crewSettingsButton;

    @FXML
    private TextField sourceWorkbookField;

    @FXML
    private ComboBox<DaySheetOption> sheetComboBox;

    @FXML
    private RangeSlider hoursRangeSlider;

    @FXML
    private TextField targetRegisterCountField;

    @FXML
    private Label hoursLabel;

    @FXML
    private Label targetRegisterCountLabel;

    @FXML
    private Label selectedConfigLabel;

    @FXML
    private Label openersValueLabel;

    @FXML
    private Label closersValueLabel;

    @FXML
    private Label maximumAvailableCrewValueLabel;

    @FXML
    private Label maximumAvailableHourValueLabel;

    @FXML
    private Label minimumAvailableCrewValueLabel;

    @FXML
    private Label minimumAvailableHourValueLabel;

    @FXML
    private HBox writeLogActionBox;

    @FXML
    private javafx.scene.control.Button writeLogButton;

    @FXML
    private javafx.scene.control.Button clearRegistersButton;

    @FXML
    private javafx.scene.control.Button saveWorkbookButton;

    @FXML
    private VBox writeLogStatusBox;

    @FXML
    private StackPane spreadsheetHost;

    @FXML
    private StackPane accountViewPane;

    @FXML
    private Label accountViewTitleLabel;

    @FXML
    private Label accountViewSubtitleLabel;

    @FXML
    private TextField accountUsernameField;

    @FXML
    private PasswordField accountPasswordField;

    @FXML
    private Label accountConfirmPasswordLabel;

    @FXML
    private PasswordField accountConfirmPasswordField;

    @FXML
    private Label accountViewValidationLabel;

    @FXML
    private Button accountViewSubmitButton;

    @FXML
    private TextField crewSettingsSearchField;

    @FXML
    private Label crewSettingsSearchEmptyLabel;

    @FXML
    private ScrollPane crewSettingsScrollPane;

    @FXML
    private VBox crewSettingsListBox;

    private final List<CrewSettingsRowEntry> crewSettingsRowEntries = new ArrayList<>();
    private AccountViewMode activeAccountViewMode;

    @FXML
    private void initialize() {
        resetSheetSelection();
        configureHourSlider();
        configureRegisterCountField();
        refreshAccountUi();
        crewSettingsSearchField.textProperty().addListener((observable, oldValue, newValue) -> applyCrewSettingsFilter());

        sheetComboBox.getSelectionModel().selectedItemProperty().addListener((observable, oldValue, newValue) -> {
            if (newValue == null) {
                setControlsEnabled(false);
                clearSpreadsheetView();
                resetAvailabilitySummary();
                resetWriteLogStatus();
                return;
            }

            setControlsEnabled(true);
            resetWriteLogStatus();
            refreshSelectedConfigLabel();
            loadSpreadsheetView(newValue);
        });
    }

    @FXML
    private void showLoginDialog() {
        toggleAccountView(AccountViewMode.LOGIN);
    }

    @FXML
    private void showCreateAccountDialog() {
        toggleAccountView(AccountViewMode.CREATE_ACCOUNT);
    }

    @FXML
    private void logout() {
        hideAccountView();
        hideCrewSettingsView();
        currentAccount = null;
        refreshAccountUi();
    }

    @FXML
    private void toggleCrewSettings() {
        if (currentAccount == null) {
            return;
        }

        if (crewSettingsScrollPane.isVisible()) {
            hideCrewSettingsView();
            return;
        }

        try {
            populateCrewSettingsList();
            crewSettingsScrollPane.setVisible(true);
            refreshCenterPanelVisibility();
        } catch (IOException | RuntimeException exception) {
            showError("Unable to load crew settings", exception.getMessage());
        }
    }

    @FXML
    private void chooseWorkbook() {
        FileChooser chooser = createExcelChooser();
        Path selectedPath = showOpenDialog(chooser);
        if (selectedPath == null) {
            return;
        }

        sourceWorkbookField.setText(selectedPath.toString());
        loadedWorkbookPath = selectedPath;
        loadWorkbook(selectedPath);
    }

    private void loadWorkbook(Path workbookPath) {
        resetSheetSelection();
        closeLoadedWorkbook();
        registerSheetStates.clear();

        Workbook workbook = null;

        try {
            try (InputStream inputStream = Files.newInputStream(workbookPath)) {
                workbook = WorkbookFactory.create(inputStream);
            }

            List<DaySheetOption> daySheetOptions = workbookInspectorService.loadDaySheetOptions(workbook);
            loadedWorkbook = workbook;
            sheetComboBox.getItems().setAll(daySheetOptions);
            sheetComboBox.setDisable(false);
            sheetComboBox.setPromptText("Select a day");
        } catch (IOException | IllegalArgumentException exception) {
            loadedWorkbookPath = null;
            registerAssignmentService.discardWorkbook(workbook);
            closeWorkbookQuietly(workbook);
            showError("Unable to load log", exception.getMessage());
        }
    }

    private void refreshAccountUi() {
        boolean hasAccount = currentAccount != null;
        accountActionBox.setManaged(!hasAccount);
        accountActionBox.setVisible(!hasAccount);
        loginButton.setManaged(!hasAccount);
        loginButton.setVisible(!hasAccount);
        createAccountButton.setManaged(!hasAccount);
        createAccountButton.setVisible(!hasAccount);
        logoutButton.setManaged(hasAccount);
        logoutButton.setVisible(hasAccount);
        logoutButton.setDisable(!hasAccount);
        crewSettingsButton.setManaged(hasAccount);
        crewSettingsButton.setVisible(hasAccount);
        crewSettingsButton.setDisable(!hasAccount);

        if (!hasAccount) {
            accountStatusLabel.setText("Guest mode. You can use Log Writer without signing in.");
            accountStatusLabel.getStyleClass().setAll("hint-label");
            return;
        }

        accountStatusLabel.setText("Logged in as " + currentAccount.username());
        accountStatusLabel.getStyleClass().setAll("hint-label", "status-success-label");
    }

    private void resetSheetSelection() {
        cancelPendingRegisterToggle();
        recentRegisterToggle = null;
        sheetComboBox.getItems().clear();
        sheetComboBox.getSelectionModel().clearSelection();
        sheetComboBox.setDisable(true);
        sheetComboBox.setPromptText("Load a log first");
        setControlsEnabled(false);
        clearSpreadsheetView();
        resetAvailabilitySummary();
        resetWriteLogStatus();
    }

    private void configureHourSlider() {
        hoursRangeSlider.setMin(8);
        hoursRangeSlider.setMax(20);
        hoursRangeSlider.setLowValue(8);
        hoursRangeSlider.setHighValue(20);
        hoursRangeSlider.setMajorTickUnit(1);
        hoursRangeSlider.setMinorTickCount(0);
        hoursRangeSlider.setBlockIncrement(1);
        hoursRangeSlider.setSnapToTicks(true);
        hoursRangeSlider.setShowTickMarks(true);
        hoursRangeSlider.setShowTickLabels(true);
        hoursRangeSlider.setLabelFormatter(new StringConverter<>() {
            @Override
            public String toString(Number value) {
                return formatHour(value.intValue());
            }

            @Override
            public Number fromString(String string) {
                return 8;
            }
        });

        hoursRangeSlider.lowValueProperty().addListener((observable, oldValue, newValue) -> refreshSelectedConfigLabel());
        hoursRangeSlider.lowValueProperty().addListener((observable, oldValue, newValue) -> resetWriteLogStatus());
        hoursRangeSlider.lowValueProperty().addListener((observable, oldValue, newValue) -> refreshAvailabilitySummaryForSelectedSheet());
        hoursRangeSlider.highValueProperty().addListener((observable, oldValue, newValue) -> refreshSelectedConfigLabel());
        hoursRangeSlider.highValueProperty().addListener((observable, oldValue, newValue) -> resetWriteLogStatus());
        hoursRangeSlider.highValueProperty().addListener((observable, oldValue, newValue) -> refreshAvailabilitySummaryForSelectedSheet());
    }

    private void configureRegisterCountField() {
        TextFormatter<String> integerFormatter = new TextFormatter<>(change -> {
            String newText = change.getControlNewText();
            if (newText.matches("\\d*")) {
                return change;
            }
            return null;
        });
        targetRegisterCountField.setTextFormatter(integerFormatter);
        targetRegisterCountField.textProperty().addListener((observable, oldValue, newValue) -> refreshSelectedConfigLabel());
        targetRegisterCountField.textProperty().addListener((observable, oldValue, newValue) -> resetWriteLogStatus());
        targetRegisterCountField.textProperty().addListener((observable, oldValue, newValue) -> refreshWriteLogButtonState());
    }

    private void setControlsEnabled(boolean enabled) {
        hoursLabel.setDisable(!enabled);
        hoursRangeSlider.setDisable(!enabled);
        targetRegisterCountLabel.setDisable(!enabled);
        targetRegisterCountField.setDisable(!enabled);
        writeLogActionBox.setManaged(enabled);
        writeLogActionBox.setVisible(enabled);

        if (!enabled) {
            hoursRangeSlider.setLowValue(8);
            hoursRangeSlider.setHighValue(20);
            targetRegisterCountField.clear();
            selectedConfigLabel.setText("Select a day to unlock the hour range, register count, and write action.");
            writeLogButton.setDisable(true);
            resetWriteLogStatus();
            return;
        }

        refreshSelectedConfigLabel();
        refreshWriteLogButtonState();
    }

    private void populateCrewSettingsList() throws IOException {
        if (currentAccount == null) {
            crewSettingsRowEntries.clear();
            crewSettingsListBox.getChildren().clear();
            return;
        }

        List<String> filteredCrewMemberNames = crewMemberDirectoryService.loadCrewMemberNamesExcludingMates();
        Map<String, Integer> savedFloorPriorities = crewFloorPriorityService.loadFloorPriorities(currentAccount);

        crewSettingsRowEntries.clear();
        crewSettingsListBox.getChildren().clear();
        crewSettingsSearchField.clear();
        if (filteredCrewMemberNames.isEmpty()) {
            Label emptyStateLabel = new Label("No crew members were found in the current data files.");
            emptyStateLabel.getStyleClass().add("hint-label");
            crewSettingsListBox.getChildren().add(emptyStateLabel);
            crewSettingsSearchEmptyLabel.setManaged(false);
            crewSettingsSearchEmptyLabel.setVisible(false);
            return;
        }

        for (String crewMemberName : filteredCrewMemberNames) {
            HBox row = createCrewSettingsRow(
                    crewMemberName,
                    savedFloorPriorities.getOrDefault(crewFloorPriorityService.normalizeName(crewMemberName), 0)
            );
            crewSettingsRowEntries.add(new CrewSettingsRowEntry(
                    crewFloorPriorityService.normalizeName(crewMemberName),
                    row
            ));
            crewSettingsListBox.getChildren().add(row);
        }
        applyCrewSettingsFilter();
    }

    private HBox createCrewSettingsRow(String crewMemberName, int initialFloorPriority) {
        Label nameLabel = new Label(crewMemberName);
        nameLabel.getStyleClass().add("crew-settings-name-label");
        nameLabel.setWrapText(true);
        nameLabel.setMinWidth(220);
        nameLabel.setPrefWidth(260);

        Label sliderLabel = new Label("Floor Priority");
        sliderLabel.getStyleClass().add("crew-settings-slider-label");

        Slider floorPrioritySlider = createFloorPrioritySlider(initialFloorPriority);
        LocalAccount account = currentAccount;
        int[] savedFloorPriority = {initialFloorPriority};
        Runnable commitFloorPriority = () -> saveFloorPrioritySelection(
                floorPrioritySlider,
                account,
                crewMemberName,
                savedFloorPriority
        );
        floorPrioritySlider.valueChangingProperty().addListener((observable, wasChanging, isChanging) -> {
            if (!isChanging) {
                commitFloorPriority.run();
            }
        });
        floorPrioritySlider.setOnMouseReleased(event -> commitFloorPriority.run());
        floorPrioritySlider.setOnKeyReleased(event -> commitFloorPriority.run());

        VBox sliderBox = new VBox(6, sliderLabel, floorPrioritySlider);
        sliderBox.getStyleClass().add("crew-settings-slider-box");
        HBox.setHgrow(sliderBox, Priority.ALWAYS);

        HBox row = new HBox(18, nameLabel, sliderBox);
        row.setAlignment(Pos.CENTER_LEFT);
        row.getStyleClass().add("crew-settings-row");
        return row;
    }

    private Slider createFloorPrioritySlider(int initialFloorPriority) {
        Slider floorPrioritySlider = new Slider(0, 3, initialFloorPriority);
        floorPrioritySlider.setMajorTickUnit(1);
        floorPrioritySlider.setMinorTickCount(0);
        floorPrioritySlider.setBlockIncrement(1);
        floorPrioritySlider.setSnapToTicks(true);
        floorPrioritySlider.setShowTickMarks(true);
        floorPrioritySlider.setShowTickLabels(true);
        floorPrioritySlider.setMaxWidth(Double.MAX_VALUE);
        return floorPrioritySlider;
    }

    private void saveFloorPrioritySelection(
            Slider floorPrioritySlider,
            LocalAccount account,
            String crewMemberName,
            int[] savedFloorPriority
    ) {
        int roundedPriority = normalizeFloorPriority(floorPrioritySlider.getValue());
        if (Math.abs(floorPrioritySlider.getValue() - roundedPriority) > 0.001) {
            floorPrioritySlider.setValue(roundedPriority);
        }

        if (roundedPriority == savedFloorPriority[0]) {
            return;
        }

        int previousPriority = savedFloorPriority[0];
        try {
            crewFloorPriorityService.saveFloorPriority(account, crewMemberName, roundedPriority);
            savedFloorPriority[0] = roundedPriority;
        } catch (IOException exception) {
            floorPrioritySlider.setValue(previousPriority);
            showError("Unable to save crew settings", exception.getMessage());
        }
    }

    private int normalizeFloorPriority(double floorPriority) {
        return Math.max(0, Math.min(3, (int) Math.round(floorPriority)));
    }

    private void applyCrewSettingsFilter() {
        if (crewSettingsSearchField == null || crewSettingsSearchEmptyLabel == null) {
            return;
        }

        String normalizedQuery = crewFloorPriorityService.normalizeName(crewSettingsSearchField.getText());
        boolean anyVisible = false;
        for (CrewSettingsRowEntry rowEntry : crewSettingsRowEntries) {
            boolean matches = normalizedQuery.isBlank() || rowEntry.normalizedCrewMemberName().contains(normalizedQuery);
            rowEntry.row().setManaged(matches);
            rowEntry.row().setVisible(matches);
            if (matches) {
                anyVisible = true;
            }
        }

        boolean showSearchEmptyState = !crewSettingsRowEntries.isEmpty() && !anyVisible;
        crewSettingsSearchEmptyLabel.setManaged(showSearchEmptyState);
        crewSettingsSearchEmptyLabel.setVisible(showSearchEmptyState);
    }

    private void toggleAccountView(AccountViewMode accountViewMode) {
        if (currentAccount != null) {
            return;
        }

        if (accountViewPane.isVisible() && activeAccountViewMode == accountViewMode) {
            hideAccountView();
            return;
        }

        configureAccountView(accountViewMode);
        accountViewPane.setVisible(true);
        refreshCenterPanelVisibility();
        javafx.application.Platform.runLater(accountUsernameField::requestFocus);
    }

    private void configureAccountView(AccountViewMode accountViewMode) {
        activeAccountViewMode = accountViewMode;
        accountViewTitleLabel.setText(accountViewMode.title());
        accountViewSubtitleLabel.setText(accountViewMode.subtitle());
        accountViewSubmitButton.setText(accountViewMode.submitButtonText());

        boolean showConfirmationField = accountViewMode.requiresPasswordConfirmation();
        accountConfirmPasswordLabel.setManaged(showConfirmationField);
        accountConfirmPasswordLabel.setVisible(showConfirmationField);
        accountConfirmPasswordField.setManaged(showConfirmationField);
        accountConfirmPasswordField.setVisible(showConfirmationField);

        clearAccountViewFields();
        clearAccountValidationMessage();
    }

    @FXML
    private void submitAccountView() {
        if (activeAccountViewMode == null) {
            return;
        }

        clearAccountValidationMessage();

        String username = accountUsernameField.getText();
        String password = accountPasswordField.getText();
        if ((username == null || username.isBlank()) || (password == null || password.isBlank())) {
            setAccountValidationMessage("Enter both a username and password.");
            return;
        }

        if (activeAccountViewMode.requiresPasswordConfirmation()
                && !Objects.equals(password, accountConfirmPasswordField.getText())) {
            setAccountValidationMessage("The passwords do not match.");
            return;
        }

        try {
            LocalAccount localAccount = activeAccountViewMode == AccountViewMode.LOGIN
                    ? localAccountService.authenticate(username, password)
                    : localAccountService.createAccount(username, password);
            currentAccount = localAccount;
            hideAccountView();
            refreshAccountUi();
        } catch (IllegalArgumentException | IOException exception) {
            setAccountValidationMessage(exception.getMessage());
        }
    }

    @FXML
    private void hideAccountView() {
        activeAccountViewMode = null;
        accountViewPane.setVisible(false);
        clearAccountViewFields();
        clearAccountValidationMessage();
        refreshCenterPanelVisibility();
    }

    private void clearAccountViewFields() {
        if (accountUsernameField != null) {
            accountUsernameField.clear();
        }
        if (accountPasswordField != null) {
            accountPasswordField.clear();
        }
        if (accountConfirmPasswordField != null) {
            accountConfirmPasswordField.clear();
        }
    }

    private void clearAccountValidationMessage() {
        if (accountViewValidationLabel == null) {
            return;
        }

        accountViewValidationLabel.setText("");
        accountViewValidationLabel.setManaged(false);
        accountViewValidationLabel.setVisible(false);
    }

    private void setAccountValidationMessage(String message) {
        accountViewValidationLabel.getStyleClass().setAll("hint-label", "status-warning-label");
        accountViewValidationLabel.setText(message);
        accountViewValidationLabel.setManaged(true);
        accountViewValidationLabel.setVisible(true);
    }

    private void hideCrewSettingsView() {
        crewSettingsScrollPane.setVisible(false);
        refreshCenterPanelVisibility();
    }

    private void refreshCenterPanelVisibility() {
        boolean showAccountView = accountViewPane.isVisible();
        accountViewPane.setManaged(showAccountView);

        boolean showCrewSettings = crewSettingsScrollPane.isVisible();
        crewSettingsScrollPane.setManaged(showCrewSettings);

        boolean showSpreadsheet = !showAccountView && !showCrewSettings && !spreadsheetHost.getChildren().isEmpty();
        spreadsheetHost.setManaged(showSpreadsheet);
        spreadsheetHost.setVisible(showSpreadsheet);
    }

    private void refreshSelectedConfigLabel() {
        if (sheetComboBox.getValue() == null) {
            return;
        }

        String registerCount = targetRegisterCountField.getText().isBlank()
                ? "not set"
                : targetRegisterCountField.getText();

        selectedConfigLabel.setText(
                "Ready for %s | Hours: %s to %s | Target Register Count: %s".formatted(
                        sheetComboBox.getValue().getDayLabel(),
                        formatHour((int) hoursRangeSlider.getLowValue()),
                        formatHour((int) hoursRangeSlider.getHighValue()),
                        registerCount
                )
        );
    }

    private void refreshWriteLogButtonState() {
        boolean hasSelectedSheet = sheetComboBox.getValue() != null;
        writeLogButton.setDisable(!hasSelectedSheet || !hasPositiveTargetRegisterCount());
    }

    private boolean hasPositiveTargetRegisterCount() {
        String targetRegisterCount = targetRegisterCountField.getText();
        if (targetRegisterCount == null || targetRegisterCount.isBlank()) {
            return false;
        }

        try {
            return Integer.parseInt(targetRegisterCount) > 0;
        } catch (NumberFormatException exception) {
            return false;
        }
    }

    private void loadSpreadsheetView(DaySheetOption selectedSheet) {
        cancelPendingRegisterToggle();
        recentRegisterToggle = null;
        if (loadedWorkbook == null) {
            clearSpreadsheetView();
            return;
        }

        try {
            SpreadsheetView spreadsheetView = workbookSpreadsheetViewService.createSpreadsheetView(
                    loadedWorkbook,
                    selectedSheet.getSheetIndex()
            );
            configureSpreadsheetViewInteractions(spreadsheetView, selectedSheet.getSheetIndex());
            spreadsheetView.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
            spreadsheetHost.getChildren().setAll(spreadsheetView);
            refreshCenterPanelVisibility();
            refreshAvailabilitySummary(selectedSheet.getSheetIndex());
            refreshYellowActionVisibility(selectedSheet);
        } catch (RuntimeException exception) {
            clearSpreadsheetView();
            resetAvailabilitySummary();
            setSaveButtonVisible(false);
            setClearRegistersButtonVisible(false);
            showError("Unable to render sheet", exception.getMessage());
        }
    }

    private void configureSpreadsheetViewInteractions(SpreadsheetView spreadsheetView, int sheetIndex) {
        spreadsheetView.getGrid().addEventHandler(
                GridChange.GRID_CHANGE_EVENT,
                gridChange -> handleSpreadsheetCellEditCommit(spreadsheetView, sheetIndex, gridChange)
        );
        spreadsheetView.addEventFilter(KeyEvent.KEY_PRESSED, keyEvent -> handleSpreadsheetKeyPress(spreadsheetView, keyEvent));
        spreadsheetView.editingCellProperty().addListener(
                (observable, oldValue, newValue) -> handleEditingCellChanged(spreadsheetView, oldValue, newValue)
        );
        spreadsheetView.addEventFilter(MouseEvent.MOUSE_RELEASED, mouseEvent -> {
            if (mouseEvent.getButton() != MouseButton.PRIMARY) {
                return;
            }

            int clickCount = mouseEvent.getClickCount();
            javafx.application.Platform.runLater(() -> handleSpreadsheetPrimaryClick(spreadsheetView, sheetIndex, clickCount));
        });
    }

    private void handleSpreadsheetPrimaryClick(SpreadsheetView spreadsheetView, int sheetIndex, int clickCount) {
        SpreadsheetViewSelectionModel selectionModel = spreadsheetView.getSelectionModel();
        var focusedCell = selectionModel.getFocusedCell();
        if (focusedCell == null) {
            return;
        }

        int rowIndex = spreadsheetView.getModelRow(focusedCell.getRow());
        int columnIndex = spreadsheetView.getModelColumn(focusedCell.getColumn());
        if (rowIndex < 0 || columnIndex < 0) {
            return;
        }

        if (isCrewNameActionColumn(columnIndex)) {
            if (clickCount == 1) {
                toggleCrewRowXFill(spreadsheetView, sheetIndex, rowIndex);
            }
            return;
        }

        if (clickCount >= 2) {
            cancelPendingRegisterToggle();
            revertRecentRegisterToggleForInlineEdit(spreadsheetView, sheetIndex, rowIndex, columnIndex);
            startInlineCellEdit(spreadsheetView, focusedCell.getRow(), focusedCell.getColumn());
            return;
        }

        scheduleRegisterToggle(spreadsheetView, sheetIndex, rowIndex, columnIndex);
    }

    private boolean isCrewNameActionColumn(int columnIndex) {
        return columnIndex == 0;
    }

    private void toggleCrewRowXFill(SpreadsheetView spreadsheetView, int sheetIndex, int rowIndex) {
        if (loadedWorkbook == null) {
            return;
        }

        DaySheetOption selectedSheet = sheetComboBox.getValue();
        if (selectedSheet == null || selectedSheet.getSheetIndex() != sheetIndex) {
            return;
        }

        try {
            flushPendingRegisterToggle();
            recentRegisterToggle = null;

            Sheet sheet = loadedWorkbook.getSheetAt(sheetIndex);
            RegisterSheetState registerSheetState = getOrCreateRegisterSheetState(sheetIndex);
            boolean changed = registerAssignmentService.toggleCrewRowXFill(
                    loadedWorkbook,
                    sheet,
                    registerSheetState,
                    rowIndex
            );
            if (!changed) {
                return;
            }

            loadSpreadsheetView(selectedSheet);
            resetWriteLogStatus();
        } catch (RuntimeException exception) {
            showError("Unable to toggle crew row", exception.getMessage());
        }
    }

    private void handleSpreadsheetKeyPress(SpreadsheetView spreadsheetView, KeyEvent keyEvent) {
        if (keyEvent.getCode() != KeyCode.ENTER) {
            return;
        }

        TablePosition<?, ?> editingCell = spreadsheetView.getEditingCell();
        if (editingCell == null) {
            return;
        }

        pendingInlineEditAdvance = new PendingInlineEditAdvance(
                spreadsheetView,
                spreadsheetView.getModelRow(editingCell.getRow()),
                spreadsheetView.getModelColumn(editingCell.getColumn())
        );
    }

    private void handleEditingCellChanged(
            SpreadsheetView spreadsheetView,
            TablePosition<?, ?> oldValue,
            TablePosition<?, ?> newValue
    ) {
        if (oldValue == null || newValue != null) {
            return;
        }
        if (pendingInlineEditAdvance == null || pendingInlineEditAdvance.spreadsheetView() != spreadsheetView) {
            return;
        }

        PendingInlineEditAdvance inlineEditAdvance = pendingInlineEditAdvance;
        pendingInlineEditAdvance = null;
        javafx.application.Platform.runLater(() -> moveSelectionToRight(
                spreadsheetView,
                inlineEditAdvance.modelRowIndex(),
                inlineEditAdvance.modelColumnIndex()
        ));
    }

    private void startInlineCellEdit(SpreadsheetView spreadsheetView, int viewRowIndex, int viewColumnIndex) {
        if (viewRowIndex < 0
                || viewColumnIndex < 0
                || viewColumnIndex >= spreadsheetView.getColumns().size()) {
            return;
        }

        spreadsheetView.requestFocus();
        spreadsheetView.edit(viewRowIndex, spreadsheetView.getColumns().get(viewColumnIndex));
    }

    private void scheduleRegisterToggle(
            SpreadsheetView spreadsheetView,
            int sheetIndex,
            int rowIndex,
            int columnIndex
    ) {
        cancelPendingRegisterToggle();
        if (!RegisterGridLayout.isHourGridCoordinate(rowIndex, columnIndex)) {
            return;
        }

        pendingRegisterToggleRequest = new PendingRegisterToggle(
                spreadsheetView,
                sheetIndex,
                rowIndex,
                columnIndex
        );
        pendingRegisterToggle = new PauseTransition(REGISTER_TOGGLE_DELAY);
        pendingRegisterToggle.setOnFinished(event -> flushPendingRegisterToggle());
        if (REGISTER_TOGGLE_DELAY.lessThanOrEqualTo(Duration.ZERO)) {
            flushPendingRegisterToggle();
            return;
        }

        pendingRegisterToggle.playFromStart();
    }

    private void flushPendingRegisterToggle() {
        PendingRegisterToggle request = pendingRegisterToggleRequest;
        cancelPendingRegisterToggle();
        if (request == null || !spreadsheetHost.getChildren().contains(request.spreadsheetView())) {
            return;
        }

        toggleManualRegisterAssignment(
                request.spreadsheetView(),
                request.sheetIndex(),
                request.rowIndex(),
                request.columnIndex()
        );
    }

    private void cancelPendingRegisterToggle() {
        if (pendingRegisterToggle != null) {
            pendingRegisterToggle.stop();
            pendingRegisterToggle = null;
        }
        pendingRegisterToggleRequest = null;
        pendingInlineEditAdvance = null;
    }

    private void revertRecentRegisterToggleForInlineEdit(
            SpreadsheetView spreadsheetView,
            int sheetIndex,
            int rowIndex,
            int columnIndex
    ) {
        if (recentRegisterToggle == null) {
            return;
        }
        if (recentRegisterToggle.spreadsheetView() != spreadsheetView
                || recentRegisterToggle.sheetIndex() != sheetIndex
                || recentRegisterToggle.rowIndex() != rowIndex
                || recentRegisterToggle.firstColumnIndex() != RegisterGridLayout.normalizeHourBlockStartColumn(columnIndex)) {
            return;
        }

        if (java.time.Instant.now().isAfter(recentRegisterToggle.toggledAt().plusMillis((long) INLINE_EDIT_REVERT_WINDOW.toMillis()))) {
            return;
        }

        toggleManualRegisterAssignment(
                spreadsheetView,
                sheetIndex,
                rowIndex,
                columnIndex,
                false
        );
        recentRegisterToggle = null;
    }

    private void handleSpreadsheetCellEditCommit(
            SpreadsheetView spreadsheetView,
            int sheetIndex,
            GridChange gridChange
    ) {
        if (loadedWorkbook == null) {
            return;
        }

        DaySheetOption selectedSheet = sheetComboBox.getValue();
        if (selectedSheet == null || selectedSheet.getSheetIndex() != sheetIndex) {
            return;
        }

        try {
            Sheet sheet = loadedWorkbook.getSheetAt(sheetIndex);
            String oldValue = Objects.toString(gridChange.getOldValue(), "");
            String newValue = Objects.toString(gridChange.getNewValue(), "");
            if (Objects.equals(oldValue, newValue)) {
                return;
            }

            applyCellTextEdit(sheet, spreadsheetView, gridChange.getRow(), gridChange.getColumn(), newValue);
            refreshAvailabilitySummary(sheetIndex);
            resetWriteLogStatus();
        } catch (RuntimeException exception) {
            showError("Unable to update cell text", exception.getMessage());
        }
    }

    @FXML
    private void writeLog() {
        DaySheetOption selectedSheet = sheetComboBox.getValue();
        if (loadedWorkbook == null || selectedSheet == null) {
            return;
        }

        try {
            flushPendingRegisterToggle();
            Sheet sheet = loadedWorkbook.getSheetAt(selectedSheet.getSheetIndex());
            RegisterSheetState registerSheetState = getOrCreateRegisterSheetState(selectedSheet.getSheetIndex());
            int startHour = (int) Math.round(hoursRangeSlider.getLowValue());
            int endHour = (int) Math.round(hoursRangeSlider.getHighValue());
            int targetRegisterCount = parseTargetRegisterCount();
            Map<String, Integer> floorPrioritiesByName = currentAccount == null
                    ? Map.of()
                    : crewFloorPriorityService.loadFloorPriorities(currentAccount);

            registerAssignmentService.applyAutomaticAssignments(
                    loadedWorkbook,
                    sheet,
                    registerSheetState,
                    startHour,
                    endHour,
                    targetRegisterCount,
                    floorPrioritiesByName
            );

            loadSpreadsheetView(selectedSheet);
            refreshWriteLogStatus(registerSheetState, startHour, endHour, targetRegisterCount);
            refreshYellowActionVisibility(selectedSheet);
        } catch (IOException | RuntimeException exception) {
            refreshYellowActionVisibility(selectedSheet);
            showError("Unable to assign registers", exception.getMessage());
        }
    }

    @FXML
    private void clearRegisters() {
        DaySheetOption selectedSheet = sheetComboBox.getValue();
        if (loadedWorkbook == null || selectedSheet == null) {
            return;
        }

        try {
            flushPendingRegisterToggle();
            recentRegisterToggle = null;

            int sheetIndex = selectedSheet.getSheetIndex();
            Sheet sheet = loadedWorkbook.getSheetAt(sheetIndex);
            RegisterSheetState registerSheetState = getOrCreateRegisterSheetState(sheetIndex);
            registerAssignmentService.clearYellowAssignments(
                    loadedWorkbook,
                    sheet,
                    registerSheetState
            );

            loadSpreadsheetView(selectedSheet);
            resetWriteLogStatus();
        } catch (RuntimeException exception) {
            refreshYellowActionVisibility(selectedSheet);
            showError("Unable to clear registers", exception.getMessage());
        }
    }

    @FXML
    private void saveFilledWorkbook() {
        DaySheetOption selectedSheet = sheetComboBox.getValue();
        if (loadedWorkbook == null || selectedSheet == null) {
            return;
        }

        try {
            flushPendingRegisterToggle();
            Path exportPath = chooseExportPath(selectedSheet);
            if (exportPath == null) {
                return;
            }

            if (Files.exists(exportPath) && !confirmOverwrite(exportPath)) {
                return;
            }

            filledWorkbookExportService.exportSheetToNewWorkbook(
                    loadedWorkbook,
                    selectedSheet.getSheetIndex(),
                    exportPath
            );
            showInformation("Workbook saved", "Saved filled workbook to:\n%s".formatted(exportPath));
        } catch (IOException | RuntimeException exception) {
            showError("Unable to save workbook", exception.getMessage());
        }
    }

    private void toggleManualRegisterAssignment(
            SpreadsheetView spreadsheetView,
            int sheetIndex,
            int rowIndex,
            int columnIndex
    ) {
        toggleManualRegisterAssignment(
                spreadsheetView,
                sheetIndex,
                rowIndex,
                columnIndex,
                true
        );
    }

    private void toggleManualRegisterAssignment(
            SpreadsheetView spreadsheetView,
            int sheetIndex,
            int rowIndex,
            int columnIndex,
            boolean rememberAsRecent
    ) {
        if (loadedWorkbook == null) {
            return;
        }

        DaySheetOption selectedSheet = sheetComboBox.getValue();
        if (selectedSheet == null || selectedSheet.getSheetIndex() != sheetIndex) {
            return;
        }

        try {
            Sheet sheet = loadedWorkbook.getSheetAt(sheetIndex);
            RegisterSheetState registerSheetState = getOrCreateRegisterSheetState(sheetIndex);

            boolean changed = registerAssignmentService.toggleManualAssignment(
                    loadedWorkbook,
                    sheet,
                    registerSheetState,
                    rowIndex,
                    columnIndex
            );

            if (changed) {
                workbookSpreadsheetViewService.refreshRegisterBlockDisplay(
                        spreadsheetView,
                        loadedWorkbook,
                        sheetIndex,
                        rowIndex,
                        RegisterGridLayout.normalizeHourBlockStartColumn(columnIndex)
                );
                if (rememberAsRecent) {
                    recentRegisterToggle = new RecentRegisterToggle(
                            spreadsheetView,
                            sheetIndex,
                            rowIndex,
                            RegisterGridLayout.normalizeHourBlockStartColumn(columnIndex),
                            java.time.Instant.now()
                    );
                }
                refreshAvailabilitySummary(sheetIndex);
                resetWriteLogStatus();
            }
        } catch (RuntimeException exception) {
            showError("Unable to update register assignment", exception.getMessage());
        }
    }

    private void refreshAvailabilitySummary(int sheetIndex) {
        if (loadedWorkbook == null || sheetIndex < 0 || sheetIndex >= loadedWorkbook.getNumberOfSheets()) {
            resetAvailabilitySummary();
            return;
        }

        Sheet sheet = loadedWorkbook.getSheetAt(sheetIndex);
        RegisterSheetState registerSheetState = getOrCreateRegisterSheetState(sheetIndex);
        CrewAvailabilitySummary summary = crewAvailabilityStatusService.summarizeAvailability(
                loadedWorkbook,
                sheet,
                registerSheetState,
                (int) Math.round(hoursRangeSlider.getLowValue()),
                (int) Math.round(hoursRangeSlider.getHighValue())
        );
        openersValueLabel.setText(Integer.toString(summary.openerCount()));
        closersValueLabel.setText(Integer.toString(summary.closerCount()));
        maximumAvailableCrewValueLabel.setText(Integer.toString(summary.maximumAvailableCrewCount()));
        maximumAvailableHourValueLabel.setText(formatOptionalHour(summary.maximumAvailableHour()));
        minimumAvailableCrewValueLabel.setText(Integer.toString(summary.minimumAvailableCrewCount()));
        minimumAvailableHourValueLabel.setText(formatOptionalHour(summary.minimumAvailableHour()));
    }

    private void refreshAvailabilitySummaryForSelectedSheet() {
        DaySheetOption selectedSheet = sheetComboBox.getValue();
        if (selectedSheet == null) {
            return;
        }

        refreshAvailabilitySummary(selectedSheet.getSheetIndex());
    }

    private void resetAvailabilitySummary() {
        openersValueLabel.setText("-");
        closersValueLabel.setText("-");
        maximumAvailableCrewValueLabel.setText("-");
        maximumAvailableHourValueLabel.setText("-");
        minimumAvailableCrewValueLabel.setText("-");
        minimumAvailableHourValueLabel.setText("-");
    }

    private void refreshWriteLogStatus(
            RegisterSheetState registerSheetState,
            int requestedStartHour,
            int requestedEndHour,
            int targetRegisterCount
    ) {
        int startHour = RegisterGridLayout.clampHour(Math.min(requestedStartHour, requestedEndHour));
        int endHour = RegisterGridLayout.clampHour(Math.max(requestedStartHour, requestedEndHour));
        int[] scheduledCountsByHour = new int[endHour - startHour + 1];

        accumulateScheduledRegisterCounts(registerSheetState.getManualAssignments(), startHour, endHour, scheduledCountsByHour);
        accumulateScheduledRegisterCounts(registerSheetState.getAutoAssignments(), startHour, endHour, scheduledCountsByHour);

        List<Label> statusLabels = new ArrayList<>();
        for (int hour = startHour; hour <= endHour; hour++) {
            int scheduledCount = scheduledCountsByHour[hour - startHour];
            if (scheduledCount >= targetRegisterCount) {
                continue;
            }

            statusLabels.add(createStatusLabel(
                    "Warning: %s only has %d scheduled registers.".formatted(
                            formatHourRange(hour),
                            scheduledCount
                    ),
                    "status-warning-label"
            ));
        }

        if (statusLabels.isEmpty()) {
            statusLabels.add(createStatusLabel("Log Written Successfully", "status-success-label"));
        }

        writeLogStatusBox.getChildren().setAll(statusLabels);
        writeLogStatusBox.setManaged(true);
        writeLogStatusBox.setVisible(true);
    }

    private void accumulateScheduledRegisterCounts(
            Iterable<RegisterBlockKey> assignments,
            int startHour,
            int endHour,
            int[] scheduledCountsByHour
    ) {
        for (RegisterBlockKey registerBlockKey : assignments) {
            int hour = RegisterGridLayout.firstColumnToHour(registerBlockKey.firstColumnIndex());
            if (hour < startHour || hour > endHour) {
                continue;
            }
            scheduledCountsByHour[hour - startHour]++;
        }
    }

    private Label createStatusLabel(String text, String styleClass) {
        Label statusLabel = new Label(text);
        statusLabel.getStyleClass().add(styleClass);
        statusLabel.setWrapText(true);
        return statusLabel;
    }

    private void resetWriteLogStatus() {
        writeLogStatusBox.getChildren().clear();
        writeLogStatusBox.setManaged(false);
        writeLogStatusBox.setVisible(false);
        refreshYellowActionVisibility();
    }

    private void clearSpreadsheetView() {
        cancelPendingRegisterToggle();
        recentRegisterToggle = null;
        spreadsheetHost.getChildren().clear();
        refreshCenterPanelVisibility();
    }

    private int parseTargetRegisterCount() {
        String targetRegisterCount = targetRegisterCountField.getText();
        if (targetRegisterCount == null || targetRegisterCount.isBlank()) {
            return 0;
        }
        return Integer.parseInt(targetRegisterCount);
    }

    private String formatHour(int hour) {
        if (hour == 12) {
            return "12PM";
        }
        if (hour == 0) {
            return "12AM";
        }
        if (hour < 12) {
            return hour + "AM";
        }
        return (hour - 12) + "PM";
    }

    private String formatOptionalHour(Integer hour) {
        return hour == null ? "-" : formatHour(hour);
    }

    private String formatHourRange(int startHour) {
        return "%s to %s".formatted(formatHour(startHour), formatHour(startHour + 1));
    }

    private RegisterSheetState getOrCreateRegisterSheetState(int sheetIndex) {
        if (loadedWorkbook == null) {
            throw new IllegalStateException("No workbook is currently loaded.");
        }

        return registerSheetStates.computeIfAbsent(
                sheetIndex,
                ignored -> registerAssignmentService.createSheetState(
                        loadedWorkbook,
                        loadedWorkbook.getSheetAt(sheetIndex)
                )
        );
    }

    private FileChooser createExcelChooser() {
        FileChooser chooser = new FileChooser();
        chooser.setTitle("Choose Log");
        chooser.getExtensionFilters().setAll(
                new FileChooser.ExtensionFilter("Excel Workbooks", "*.xlsx", "*.xlsm", "*.xls")
        );
        return chooser;
    }

    private Path showOpenDialog(FileChooser chooser) {
        Window owner = resolveOwnerWindow();
        var selectedFile = chooser.showOpenDialog(owner);
        return selectedFile == null ? null : selectedFile.toPath();
    }

    private DirectoryChooser createExportDirectoryChooser(String exportFileName) {
        DirectoryChooser chooser = new DirectoryChooser();
        chooser.setTitle("Choose Save Location for %s".formatted(exportFileName));
        if (loadedWorkbookPath != null && loadedWorkbookPath.getParent() != null && Files.isDirectory(loadedWorkbookPath.getParent())) {
            chooser.setInitialDirectory(loadedWorkbookPath.getParent().toFile());
        }
        return chooser;
    }

    private Path chooseExportPath(DaySheetOption selectedSheet) {
        String exportFileName = buildExportFileName(selectedSheet.getSheetName());
        DirectoryChooser chooser = createExportDirectoryChooser(exportFileName);
        Window owner = resolveOwnerWindow();
        var selectedDirectory = chooser.showDialog(owner);
        if (selectedDirectory == null) {
            return null;
        }

        return selectedDirectory.toPath().resolve(exportFileName);
    }

    private String buildExportFileName(String sheetName) {
        String safeSheetName = sanitizeFileName(sheetName);
        if (safeSheetName.isBlank()) {
            safeSheetName = "FilledWorkbook";
        }
        return safeSheetName + "_Filled" + resolveExportExtension();
    }

    private String sanitizeFileName(String fileName) {
        return fileName
                .replaceAll("[\\\\/:*?\"<>|]", "_")
                .replaceAll("[. ]+$", "")
                .trim();
    }

    private String resolveExportExtension() {
        if (loadedWorkbook == null) {
            return ".xlsx";
        }
        return filledWorkbookExportService.resolveExportExtension(loadedWorkbook);
    }

    private boolean confirmOverwrite(Path exportPath) {
        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        alert.setHeaderText("Replace existing workbook?");
        alert.setContentText("%s already exists.\nDo you want to replace it?".formatted(exportPath.getFileName()));
        Window owner = resolveOwnerWindow();
        if (owner != null) {
            alert.initOwner(owner);
        }
        return alert.showAndWait().orElse(ButtonType.CANCEL) == ButtonType.OK;
    }

    private void closeLoadedWorkbook() {
        cancelPendingRegisterToggle();
        recentRegisterToggle = null;
        registerAssignmentService.discardWorkbook(loadedWorkbook);
        closeWorkbookQuietly(loadedWorkbook);
        loadedWorkbook = null;
    }

    private void closeWorkbookQuietly(Workbook workbook) {
        if (workbook == null) {
            return;
        }

        try {
            workbook.close();
        } catch (IOException ignored) {
        }
    }

    private void showError(String headerText, String message) {
        Alert alert = new Alert(Alert.AlertType.ERROR);
        alert.setHeaderText(headerText);
        alert.setContentText(message);
        Window owner = resolveOwnerWindow();
        if (owner != null) {
            alert.initOwner(owner);
        }
        alert.showAndWait();
    }

    private void showInformation(String headerText, String message) {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setHeaderText(headerText);
        alert.setContentText(message);
        Window owner = resolveOwnerWindow();
        if (owner != null) {
            alert.initOwner(owner);
        }
        alert.showAndWait();
    }

    private Window resolveOwnerWindow() {
        return sourceWorkbookField.getScene() == null ? null : sourceWorkbookField.getScene().getWindow();
    }

    private void setSaveButtonVisible(boolean visible) {
        saveWorkbookButton.setManaged(visible);
        saveWorkbookButton.setVisible(visible);
        saveWorkbookButton.setDisable(!visible);
    }

    private void setClearRegistersButtonVisible(boolean visible) {
        clearRegistersButton.setManaged(visible);
        clearRegistersButton.setVisible(visible);
        clearRegistersButton.setDisable(!visible);
    }

    private void refreshYellowActionVisibility() {
        refreshYellowActionVisibility(sheetComboBox.getValue());
    }

    private void refreshYellowActionVisibility(DaySheetOption selectedSheet) {
        if (loadedWorkbook == null || selectedSheet == null) {
            setSaveButtonVisible(false);
            setClearRegistersButtonVisible(false);
            return;
        }

        int sheetIndex = selectedSheet.getSheetIndex();
        if (sheetIndex < 0 || sheetIndex >= loadedWorkbook.getNumberOfSheets()) {
            setSaveButtonVisible(false);
            setClearRegistersButtonVisible(false);
            return;
        }

        boolean hasYellowAssignments = registerAssignmentService.hasYellowAssignments(loadedWorkbook.getSheetAt(sheetIndex));
        setSaveButtonVisible(hasYellowAssignments);
        setClearRegistersButtonVisible(hasYellowAssignments);
    }

    private void moveSelectionToRight(SpreadsheetView spreadsheetView, int modelRowIndex, int modelColumnIndex) {
        if (modelRowIndex < 0
                || modelRowIndex >= spreadsheetView.getGrid().getRowCount()
                || modelColumnIndex < 0
                || modelColumnIndex >= spreadsheetView.getGrid().getColumnCount()) {
            return;
        }

        var rowCells = spreadsheetView.getGrid().getRows().get(modelRowIndex);
        if (modelColumnIndex >= rowCells.size()) {
            return;
        }

        var currentCell = rowCells.get(modelColumnIndex);
        if (currentCell == null) {
            return;
        }

        SpreadsheetViewSelectionModel selectionModel = spreadsheetView.getSelectionModel();
        selectionModel.clearAndSelect(currentCell);
        selectionModel.clearAndSelectRightCell();
    }

    private void applyCellTextEdit(
            Sheet sheet,
            SpreadsheetView spreadsheetView,
            int rowIndex,
            int columnIndex,
            String value
    ) {
        CellRangeAddress mergedRegion = findMergedRegionContaining(sheet, rowIndex, columnIndex);
        if (mergedRegion != null) {
            setCellText(sheet, mergedRegion.getFirstRow(), mergedRegion.getFirstColumn(), value);
            return;
        }

        if (shouldMirrorHourPairEdit(spreadsheetView, rowIndex, columnIndex)) {
            int firstColumnIndex = RegisterGridLayout.normalizeHourBlockStartColumn(columnIndex);
            setCellText(sheet, rowIndex, firstColumnIndex, value);
            setCellText(sheet, rowIndex, firstColumnIndex + 1, value);
            return;
        }

        setCellText(sheet, rowIndex, columnIndex, value);
    }

    private boolean shouldMirrorHourPairEdit(
            SpreadsheetView spreadsheetView,
            int rowIndex,
            int columnIndex
    ) {
        if (!RegisterGridLayout.isHourGridCoordinate(rowIndex, columnIndex)) {
            return false;
        }

        int firstColumnIndex = RegisterGridLayout.normalizeHourBlockStartColumn(columnIndex);
        int secondColumnIndex = firstColumnIndex + 1;
        if (rowIndex >= spreadsheetView.getGrid().getRowCount()
                || secondColumnIndex >= spreadsheetView.getGrid().getColumnCount()) {
            return false;
        }

        return spreadsheetView.getGrid().getRows().get(rowIndex).get(firstColumnIndex)
                == spreadsheetView.getGrid().getRows().get(rowIndex).get(secondColumnIndex);
    }

    private CellRangeAddress findMergedRegionContaining(Sheet sheet, int rowIndex, int columnIndex) {
        for (int mergedRegionIndex = 0; mergedRegionIndex < sheet.getNumMergedRegions(); mergedRegionIndex++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(mergedRegionIndex);
            if (mergedRegion.isInRange(rowIndex, columnIndex)) {
                return mergedRegion;
            }
        }
        return null;
    }

    private void setCellText(Sheet sheet, int rowIndex, int columnIndex, String value) {
        Cell cell = getOrCreateCell(sheet, rowIndex, columnIndex);
        if (value == null || value.isBlank()) {
            cell.setBlank();
            return;
        }

        cell.setCellValue(value);
    }

    private Cell getCell(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return null;
        }

        return row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
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

    private record PendingRegisterToggle(
            SpreadsheetView spreadsheetView,
            int sheetIndex,
            int rowIndex,
            int columnIndex
    ) {
    }

    private record PendingInlineEditAdvance(
            SpreadsheetView spreadsheetView,
            int modelRowIndex,
            int modelColumnIndex
    ) {
    }

    private record RecentRegisterToggle(
            SpreadsheetView spreadsheetView,
            int sheetIndex,
            int rowIndex,
            int firstColumnIndex,
            java.time.Instant toggledAt
    ) {
    }

    private enum AccountViewMode {
        LOGIN("Login", "Sign in to an existing local account.", "Login", false),
        CREATE_ACCOUNT("Create Account", "Create a local account stored on this device.", "Create Account", true);

        private final String title;
        private final String subtitle;
        private final String submitButtonText;
        private final boolean requiresPasswordConfirmation;

        AccountViewMode(
                String title,
                String subtitle,
                String submitButtonText,
                boolean requiresPasswordConfirmation
        ) {
            this.title = title;
            this.subtitle = subtitle;
            this.submitButtonText = submitButtonText;
            this.requiresPasswordConfirmation = requiresPasswordConfirmation;
        }

        private String title() {
            return title;
        }

        private String subtitle() {
            return subtitle;
        }

        private String submitButtonText() {
            return submitButtonText;
        }

        private boolean requiresPasswordConfirmation() {
            return requiresPasswordConfirmation;
        }
    }

    private record CrewSettingsRowEntry(
            String normalizedCrewMemberName,
            HBox row
    ) {
    }
}
