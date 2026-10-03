package io.github.r2d2c2.gitsheet;

import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.application.Application;
import javafx.application.Platform;
import javafx.beans.property.ReadOnlyStringWrapper;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.chart.*;
import javafx.scene.control.*;
import javafx.scene.input.*;
import javafx.scene.input.DataFormat;
import javafx.scene.layout.*;
import javafx.stage.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.*;
import java.util.function.Supplier;
import java.util.function.BooleanSupplier;

public final class SpreadsheetApp extends Application {
    private Book book = new Book();
    private static final DataFormat CELL_COPY = new DataFormat("application/x-git-sheet-copy-id");
    private CellClipboard cellClipboard;
    private String copyId;
    private record CutRange(int sheet, int row, int column, int rows, int columns) {}
    private CutRange pendingCut;
    private RangeFilter.Spec rangeFilter;
    private int filterOffset;
    private List<Integer> filterMatches = List.of();
    private Stage stage;
    private Path document;
    private boolean dirty;
    private int sheetIndex, rowWindow, columnWindow;
    private static final int ROW_WINDOW = 1000, COLUMN_WINDOW = 52;
    @FXML private TableView<Integer> grid;
    @FXML private TabPane tabs;
    @FXML private Pane mergeLayer;
    private MergedCellOverlay mergeOverlay;
    @FXML private TextField address, filter;
    @FXML private TextArea formula;
    @FXML private Label status, selectionInfo;
    @FXML private ComboBox<String> formats;
    private final ExecutorService background = Executors.newVirtualThreadPerTaskExecutor();
    private boolean changingTabs;
    private SaveNotification saveNotification;
    private GitDialog gitDialog;
    private boolean gitBusy;
    private BooleanSupplier pendingEdit;
    private final javafx.animation.PauseTransition rowLayoutRefresh = new javafx.animation.PauseTransition(javafx.util.Duration.millis(100));
    private record Position(int row, int column) {}
    private record Bounds(int firstRow, int lastRow, int firstCol, int lastCol) {}

    public static void main(String[] args) { launch(args); }
    @Override public void start(Stage stage) throws IOException {
        this.stage = stage;
        var loader = new FXMLLoader(Objects.requireNonNull(getClass().getResource("/spreadsheet.fxml")));
        loader.setController(this);
        BorderPane root = loader.load();
        bindActions(); configureInputs();
        mergeOverlay = new MergedCellOverlay(grid, mergeLayer, () -> book, () -> sheetIndex,
                () -> { var p = selected(); return p == null ? null : new org.apache.poi.ss.util.CellAddress(p.row(), p.column()); },
                (p, edit) -> focusMergeAnchor(new Position(p.getRow(), p.getColumn()), edit), this::cellCss);
        grid.setEditable(true); grid.setFixedCellSize(-1); grid.getSelectionModel().setCellSelectionEnabled(true);
        rowLayoutRefresh.setOnFinished(_ -> grid.refresh());
        grid.setRowFactory(_ -> new TableRow<>() {
            @Override protected void updateItem(Integer index, boolean empty) {
                super.updateItem(index, empty);
                var row = empty || index == null ? null : book.workbook().getSheetAt(sheetIndex).getRow(index);
                double height = row != null && row.getCTRow().isSetHt() ? Math.max(1, row.getHeightInPoints() * 96.0 / 72) : USE_COMPUTED_SIZE;
                setMinHeight(height); setPrefHeight(height); setMaxHeight(height);
            }
        });
        grid.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        grid.getSelectionModel().getSelectedCells().addListener((javafx.collections.ListChangeListener<TablePosition>) change -> selectionChanged());
        grid.addEventFilter(KeyEvent.KEY_PRESSED, this::gridKey);
        grid.addEventFilter(KeyEvent.KEY_TYPED, this::typeIntoCell);
        grid.setOnInputMethodTextChanged(event -> {
            if (pendingEdit != null || event.getTarget() instanceof TextInputControl
                    || (event.getCommitted().isEmpty() && event.getComposed().isEmpty())) return;
            var editor = beginTyping();
            if (editor != null) {
                editor.fireEvent(new InputMethodEvent(InputMethodEvent.INPUT_METHOD_TEXT_CHANGED,
                        event.getComposed(), event.getCommitted(), event.getCaretPosition()));
                event.consume();
            }
        });
        grid.setInputMethodRequests(new InputMethodRequests() {
            @Override public javafx.geometry.Point2D getTextLocation(int offset) {
                var point = grid.localToScreen(64, 30);
                return point == null ? new javafx.geometry.Point2D(0, 0) : point;
            }
            @Override public int getLocationOffset(int x, int y) { return 0; }
            @Override public void cancelLatestCommittedText() { }
            @Override public String getSelectedText() { return ""; }
        });
        grid.setPlaceholder(new Label("조건에 맞는 행이 없습니다."));
        tabs.setTabClosingPolicy(TabPane.TabClosingPolicy.UNAVAILABLE); tabs.setPrefHeight(40);
        tabs.setStyle("-fx-open-tab-animation: none; -fx-close-tab-animation: none;");
        tabs.getSelectionModel().selectedIndexProperty().addListener((_, _, index) -> {
            if (!changingTabs && index.intValue() >= 0) { rangeFilter = null; sheetIndex = index.intValue(); rowWindow = columnWindow = 0; refresh(); }
        });
        var overlay = new StackPane(root);
        saveNotification = new SaveNotification(overlay);
        var scene = new Scene(overlay, 1280, 820);
        scene.getStylesheets().add(Objects.requireNonNull(getClass().getResource("/theme.css")).toExternalForm());
        scene.getAccelerators().put(new KeyCodeCombination(KeyCode.S, KeyCombination.SHORTCUT_DOWN), () -> save(false));
        scene.getAccelerators().put(new KeyCodeCombination(KeyCode.O, KeyCombination.SHORTCUT_DOWN), this::open);
        scene.getAccelerators().put(new KeyCodeCombination(KeyCode.Z, KeyCombination.SHORTCUT_DOWN), this::undo);
        scene.getAccelerators().put(new KeyCodeCombination(KeyCode.Y, KeyCombination.SHORTCUT_DOWN), this::redo);
        stage.setScene(scene); stage.setMinWidth(900); stage.setMinHeight(560);
        stage.setOnCloseRequest(event -> { if (!canLeave()) event.consume(); });
        sample(); refreshTabs(); refresh(); stage.show();
        if (getParameters().getRaw().contains("--smoke-test")) {
            Platform.runLater(() -> {
                try {
                    if (grid.getColumns().size() != COLUMN_WINDOW + 1 || !book.display(0, 4, 3).equals("₩660,000"))
                        throw new IllegalStateException("UI smoke test failed: " + book.display(0, 4, 3));
                    smokeTest();
                    smokeMergedCells();
                    smokeSaveAndGit();
                } catch (Exception e) { failSmoke(e); }
            });
        }
    }
    private final Map<String, Runnable> actions = new HashMap<>();
    private void bindActions() {
        actions.put("newDocument", this::newBook);
        actions.put("openDocument", this::open);
        actions.put("saveDocument", () -> save(false));
        actions.put("saveAs", () -> save(true));
        actions.put("exportXlsx", () -> export(false));
        actions.put("exportCsv", () -> export(true));
        actions.put("undoEdit", this::undo);
        actions.put("redoEdit", this::redo);
        actions.put("copyCells", this::copy);
        actions.put("cutCells", () -> copy(true));
        actions.put("pasteCells", this::paste);
        actions.put("pasteValues", () -> paste(CellClipboard.Mode.VALUES));
        actions.put("pasteFormats", () -> paste(CellClipboard.Mode.FORMATS));
        actions.put("clearCells", this::clear);
        actions.put("findCells", this::find);
        actions.put("replaceCells", this::replaceAll);
        actions.put("fillDown", () -> fill(SheetEdits.Direction.DOWN));
        actions.put("fillRight", () -> fill(SheetEdits.Direction.RIGHT));
        actions.put("alignLeft", () -> applyStyle("align", "LEFT"));
        actions.put("alignCenter", () -> applyStyle("align", "CENTER"));
        actions.put("alignRight", () -> applyStyle("align", "RIGHT"));
        actions.put("borderAll", () -> applyStyle("border", "THIN"));
        actions.put("borderNone", () -> applyStyle("border", "NONE"));
        actions.put("fontSize", () -> prompt("글꼴 크기", "크기 (6–72)", "13").ifPresent(value -> {
                    try { int size = Integer.parseInt(value); if (size < 6 || size > 72) throw new IllegalArgumentException("6–72 범위의 크기를 입력하세요."); applyStyle("size", value); }
                    catch (Exception e) { error(e); }
                }));
        actions.put("columnWidth", this::columnWidth);
        actions.put("rowHeight", this::rowHeight);
        actions.put("autoRowHeight", () -> applyRowHeight(null));
        actions.put("mergeCells", () -> changeMerge(false));
        actions.put("unmergeCells", () -> changeMerge(true));
        actions.put("wrapOn", () -> applyStyle("wrap", "true"));
        actions.put("wrapOff", () -> applyStyle("wrap", "false"));
        actions.put("addSheet", this::addSheet);
        actions.put("renameSheet", this::renameSheet);
        actions.put("deleteSheet", this::deleteSheet);
        actions.put("insertRows", () -> structuralEdit(StructuralEdits.Axis.ROW, true));
        actions.put("deleteRows", () -> structuralEdit(StructuralEdits.Axis.ROW, false));
        actions.put("insertColumns", () -> structuralEdit(StructuralEdits.Axis.COLUMN, true));
        actions.put("deleteColumns", () -> structuralEdit(StructuralEdits.Axis.COLUMN, false));
        actions.put("previousPage", () -> page(-1));
        actions.put("nextPage", () -> page(1));
        actions.put("gitConnect", this::gitConnect);
        actions.put("gitInspect", this::gitInspect);
        actions.put("gitCommit", this::gitCommit);
        actions.put("help", () -> textDialog("Git Sheet 0.2", "JDK 25 + JavaFX\n\n"
                + "한 번 클릭 후 타이핑: 새 값 입력 · 더블 클릭 / F2: 기존 값 편집 · Enter: 적용 · Escape: 취소\n수식: =SUM(A1:A10), =IF(B1>0,\"예\",\"아니오\")\n"
                + "주소 상자: A1 또는 Z10000 입력 → 이동\n도구막대 정렬·필터는 현재 1,000행 화면에 적용됩니다. 데이터 메뉴에서 실제 범위 정렬을 사용할 수 있습니다.\n"
                + "Ctrl+C/V: 탭으로 구분된 직사각형 데이터 복사·붙여넣기\nCtrl+D/R: 아래로/오른쪽 채우기 (상대 참조 이동)\nCtrl+S: Git 친화적인 .gsheet 저장 · 저장 위치 알림 4초\n"
                + "Git → 서버 연결 / 동기화: GitHub·GitLab·Gitea·자체 서버 선택, push/fetch\n\n"
                + "지원: 값, POI 지원 수식, 기본 서식, 다중 시트, CSV/XLSX, 실행 취소, Git diff/commit, 간단한 차트\n"
                + "제한: Excel 완전 호환 제품이 아닙니다. VBA, Power Query, 피벗, 협업, 고급 차트 및 동적 배열 미지원.\n"
                + ".gsheet 변환은 그림, 이름 정의, 유효성 검사, 조건부 서식 등 고급 Excel 요소를 보존하지 않습니다.\n"
                + "병합은 표시·편집을 지원하며 틀 고정은 저장만 지원합니다. 원본 .xlsx는 따로 보관하세요.\n"
                + "외부 연결은 갱신하지 않으며 지원하지 않는 수식은 #UNSUPPORTED!로 표시됩니다."));
        actions.put("sortRange", this::sortRange);
        actions.put("filterRange", this::filterRange);
        actions.put("clearFilter", () -> { rangeFilter = null; filterOffset = 0; refreshRows(); });
        actions.put("new", this::newBook); actions.put("open", this::open); actions.put("save", () -> save(false));
        actions.put("undo", this::undo); actions.put("redo", this::redo);
        actions.put("bold", () -> toggleFont("bold")); actions.put("italic", () -> toggleFont("italic"));
        actions.put("fill", () -> applyStyle("fill", Short.toString(IndexedColors.LIGHT_YELLOW.getIndex())));
        actions.put("ascending", () -> sort(false)); actions.put("descending", () -> sort(true));
        actions.put("chart", this::chart); actions.put("git", this::gitConnect);
    }
    @FXML private void dispatchAction(javafx.event.ActionEvent event) {
        Object source = event.getSource();
        String key = (String) (source instanceof MenuItem item ? item.getUserData() : ((javafx.scene.Node) source).getUserData());
        Objects.requireNonNull(actions.get(key), key).run();
    }
    private void configureInputs() {
        formats.setItems(FXCollections.observableArrayList("일반", "숫자", "통화", "백분율", "날짜"));
        formats.setValue("일반"); formats.setOnAction(_ -> applyStyle("format", switch (formats.getValue()) {
            case "숫자" -> "#,##0.00"; case "통화" -> "₩#,##0"; case "백분율" -> "0.00%"; case "날짜" -> "yyyy-mm-dd"; default -> "General";
        }));
        filter.setOnAction(_ -> refreshRows()); address.setOnAction(_ -> goTo());
        formula.addEventFilter(KeyEvent.KEY_PRESSED, event -> {
            if (event.getCode() == KeyCode.ENTER) {
                if (event.isAltDown()) formula.replaceSelection("\n");
                else { var p = selected(); if (p != null) mutate(() -> book.set(sheetIndex, p.row(), p.column(), formula.getText())); }
                event.consume();
            } else if (event.getCode() == KeyCode.ESCAPE) { selectionChanged(); event.consume(); }
        });
    }
    private static Button button(String text, Runnable action) { var b = new Button(text); b.setOnAction(_ -> action.run()); return b; }
    private void refreshTabs() {
        changingTabs = true; tabs.getTabs().clear();
        for (var sheet : book.workbook()) tabs.getTabs().add(new Tab(sheet.getSheetName()));
        sheetIndex = Math.clamp(sheetIndex, 0, book.workbook().getNumberOfSheets() - 1);
        tabs.getSelectionModel().select(sheetIndex); changingTabs = false;
    }
    private void refresh() {
        grid.getColumns().clear();
        var numbers = new TableColumn<Integer, String>("#"); numbers.setPrefWidth(64); numbers.setSortable(false); numbers.setEditable(false); numbers.setReorderable(false);
        numbers.setCellValueFactory(data -> new ReadOnlyStringWrapper(Integer.toString(data.getValue() + 1))); grid.getColumns().add(numbers);
        for (int c = columnWindow; c < Math.min(Book.MAX_COLUMNS, columnWindow + COLUMN_WINDOW); c++) {
            final int columnIndex = c;
            var column = new TableColumn<Integer, String>(CellReference.convertNumToColString(c));
            column.setUserData(c); column.setSortable(false); column.setReorderable(false);
            column.setMinWidth(24); column.setPrefWidth(Math.max(24, book.workbook().getSheetAt(sheetIndex).getColumnWidthInPixels(c)));
            column.widthProperty().addListener((_, _, _) -> rowLayoutRefresh.playFromStart());
            column.setCellValueFactory(data -> new ReadOnlyStringWrapper(book.display(sheetIndex, data.getValue(), columnIndex)));
            column.setCellFactory(_ -> new GridCell(columnIndex));
            grid.getColumns().add(column);
        }
        refreshRows(); title();
    }
    private void refreshRows() {
        if (book.workbook().getSheetAt(sheetIndex).getNumMergedRegions() > 0) { rangeFilter = null; filter.clear(); }
        requestMergePaint();
        filter.setDisable(rangeFilter != null || book.workbook().getSheetAt(sheetIndex).getNumMergedRegions() > 0);
        if (rangeFilter != null) {
            try {
                filterMatches = RangeFilter.matchingRows(book, sheetIndex, rangeFilter);
                filterOffset = Math.min(filterOffset, Math.max(0, ((filterMatches.size() - 1) / ROW_WINDOW) * ROW_WINDOW));
                var rows = new ArrayList<Integer>();
                if (rangeFilter.header()) rows.add(rangeFilter.range().getFirstRow());
                rows.addAll(filterMatches.subList(filterOffset, Math.min(filterMatches.size(), filterOffset + ROW_WINDOW)));
                grid.setItems(FXCollections.observableArrayList(rows));
                status.setText("열별 필터 · 일치 %,d행 · 결과 페이지 %d · 시트 메뉴에서 이전/다음 1,000행".formatted(filterMatches.size(), filterOffset / ROW_WINDOW + 1));
                if (!rows.isEmpty()) grid.getSelectionModel().select(0, grid.getColumns().get(1));
                return;
            } catch (Exception e) { rangeFilter = null; filter.setDisable(false); error(e); }
        }
        filterMatches = List.of();
        String needle = filter.getText().toLowerCase(Locale.ROOT);
        var rows = new ArrayList<Integer>(ROW_WINDOW);
        for (int r = rowWindow; r < Math.min(Book.MAX_ROWS, rowWindow + ROW_WINDOW); r++) {
            boolean match = needle.isBlank();
            if (!match) for (int c = columnWindow; c < Math.min(Book.MAX_COLUMNS, columnWindow + COLUMN_WINDOW); c++) {
                if (book.display(sheetIndex, r, c).toLowerCase(Locale.ROOT).contains(needle)) { match = true; break; }
            }
            if (match) rows.add(r);
        }
        grid.setItems(FXCollections.observableArrayList(rows));
        status.setText("행 %,d–%,d · 표시 %,d행 · %s".formatted(rowWindow + 1, Math.min(Book.MAX_ROWS, rowWindow + ROW_WINDOW), rows.size(), book.workbook().getSheetName(sheetIndex)));
        if (!rows.isEmpty()) grid.getSelectionModel().select(0, grid.getColumns().get(1));
    }
    private Position selected() {
        var cells = grid.getSelectionModel().getSelectedCells();
        if (cells.isEmpty()) return null;
        var selected = cells.getLast();
        if (selected.getRow() < 0 || selected.getRow() >= grid.getItems().size() || selected.getColumn() < 1) return null;
        return mergeAnchor(grid.getItems().get(selected.getRow()), (Integer) selected.getTableColumn().getUserData());
    }
    private List<Position> selectedPositions() {
        var positions = new ArrayList<Position>();
        for (var p : grid.getSelectionModel().getSelectedCells()) {
            if (p.getColumn() > 0 && p.getRow() >= 0 && p.getRow() < grid.getItems().size())
                positions.add(new Position(grid.getItems().get(p.getRow()), (Integer) p.getTableColumn().getUserData()));
        }
        return positions;
    }
    private void changeMerge(boolean remove) {
        if (pendingEdit != null && !pendingEdit.getAsBoolean()) return;
        var b = bounds(); if (b == null || !clipboardViewIsContiguous(b)) return;
        if (rangeFilter != null || !filter.getText().isBlank()) { status.setText("필터를 해제한 뒤 병합을 변경하세요."); return; }
        var range = new org.apache.poi.ss.util.CellRangeAddress(grid.getItems().get(b.firstRow()), grid.getItems().get(b.lastRow()),
                (Integer) grid.getColumns().get(b.firstCol()).getUserData(), (Integer) grid.getColumns().get(b.lastCol()).getUserData());
        if (!remove && selectedPositions().size() != range.getNumberOfCells()) { status.setText("빈틈 없는 직사각형 범위를 선택하세요."); return; }
        if (mutate(() -> { if (remove) MergeEdits.unmerge(book, sheetIndex, range); else MergeEdits.merge(book, sheetIndex, range); })) {
            refreshRows(); focusMergeAnchor(new Position(range.getFirstRow(), range.getFirstColumn()), false);
            status.setText(remove ? "병합 해제 완료" : "병합 완료 · 좌상단 셀에서 편집합니다.");
        }
    }
    private void focusMergeAnchor(Position p, boolean edit) {
        int row = grid.getItems().indexOf(p.row());
        var column = grid.getColumns().stream().filter(c -> Objects.equals(c.getUserData(), p.column())).findFirst().orElse(null);
        if (row < 0 || column == null) {
            address.setText(Book.address(p.row(), p.column())); goTo();
            row = grid.getItems().indexOf(p.row());
            column = grid.getColumns().stream().filter(c -> Objects.equals(c.getUserData(), p.column())).findFirst().orElse(null);
        }
        if (row < 0 || column == null) return;
        grid.getSelectionModel().clearAndSelect(row, column); grid.getFocusModel().focus(row, column);
        grid.scrollTo(row); grid.scrollToColumn(column); grid.requestFocus();
        grid.applyCss(); grid.layout();
        if (edit) grid.edit(row, column);
        requestMergePaint();
    }
    private void requestMergePaint() { if (mergeOverlay != null) mergeOverlay.requestPaint(); }
    private void paintMerges() { if (mergeOverlay != null) mergeOverlay.paint(); }
    private Position mergeAnchor(int row, int column) {
        var merge = MergeEdits.containing(book.workbook().getSheetAt(sheetIndex), row, column);
        return merge == null ? new Position(row, column) : new Position(merge.getFirstRow(), merge.getFirstColumn());
    }
    private List<Position> editablePositions() {
        return selectedPositions().stream().map(p -> mergeAnchor(p.row(), p.column())).distinct().toList();
    }
    private void selectionChanged() {
        requestMergePaint();
        var p = selected(); if (p == null) return;
        address.setText(Book.address(p.row(), p.column())); formula.setText(book.input(sheetIndex, p.row(), p.column()));
        var values = editablePositions(); double sum = 0; int count = 0;
        for (var cell : values) {
            try { sum += Double.parseDouble(book.raw(sheetIndex, cell.row(), cell.column())); count++; }
            catch (NumberFormatException ignored) { }
        }
        selectionInfo.setText("선택 %d셀 · 숫자 %d · 합계 %s".formatted(values.size(), count, Double.toString(sum)));
    }
    private boolean mutate(Runnable action) {
        try { book.transaction(_ -> action.run()); pendingCut = null; dirty = true; if (rangeFilter != null) refreshRows(); else grid.refresh(); title(); selectionChanged(); return true; }
        catch (Exception e) { error(e); grid.refresh(); return false; }
    }
    private void undo() { if (book.undo()) { pendingCut = null; dirty = true; refreshTabs(); refresh(); } }
    private void redo() { if (book.redo()) { pendingCut = null; dirty = true; refreshTabs(); refresh(); } }
    private void applyStyle(String property, String value) {
        var positions = editablePositions(); if (positions.isEmpty()) return;
        mutate(() -> positions.forEach(p -> book.style(sheetIndex, p.row(), p.column(), property, value)));
    }
    private void toggleFont(String property) {
        var p = selected(); if (p == null) return;
        var cell = book.cell(sheetIndex, p.row(), p.column(), false);
        var font = book.workbook().getFontAt(cell == null ? 0 : cell.getCellStyle().getFontIndex());
        applyStyle(property, Boolean.toString(!(property.equals("bold") ? font.getBold() : font.getItalic())));
    }
    private void typeIntoCell(KeyEvent event) {
        if (event.getTarget() instanceof TextInputControl || pendingEdit != null || event.isShortcutDown()
                || event.isAltDown() || event.getCharacter().isEmpty()
                || event.getCharacter().codePoints().anyMatch(Character::isISOControl)) return;
        var editor = beginTyping();
        if (editor != null) {
            editor.setText(event.getCharacter()); editor.positionCaret(editor.getLength()); event.consume();
        }
    }
    private TextArea beginTyping() {
        var focus = grid.getFocusModel().getFocusedCell();
        if (focus.getRow() < 0 || focus.getColumn() <= 0) return null;
        var anchor = mergeAnchor(grid.getItems().get(focus.getRow()), (Integer) focus.getTableColumn().getUserData());
        focusMergeAnchor(anchor, true);
        for (var node : grid.lookupAll(".table-cell")) {
            if (node instanceof GridCell cell && cell.isEditing()) {
                cell.editor.clear(); cell.editor.applyCss(); return cell.editor;
            }
        }
        return null;
    }
    private void gridKey(KeyEvent event) {
        if (event.getTarget() instanceof TextInputControl) return;
        if (event.isShortcutDown()) {
            if (event.getCode() == KeyCode.C) { copy(); event.consume(); }
            else if (event.getCode() == KeyCode.X) { copy(true); event.consume(); }
            else if (event.getCode() == KeyCode.V) { paste(); event.consume(); }
            else if (event.getCode() == KeyCode.D) { fill(SheetEdits.Direction.DOWN); event.consume(); }
            else if (event.getCode() == KeyCode.R) { fill(SheetEdits.Direction.RIGHT); event.consume(); }
        } else if (event.getCode() == KeyCode.ESCAPE && pendingCut != null) { pendingCut = null; status.setText("잘라내기를 취소했습니다. 복사한 내용은 유지됩니다."); event.consume(); }
        else if (event.getCode() == KeyCode.DELETE || event.getCode() == KeyCode.BACK_SPACE) { clear(); event.consume(); }
        else if (event.getCode() == KeyCode.F2 || event.getCode() == KeyCode.ENTER) {
            var p = selected(); if (p != null) focusMergeAnchor(p, true); event.consume();
        }
    }
    private Bounds bounds() {
        var cells = grid.getSelectionModel().getSelectedCells().stream().filter(p -> p.getColumn() > 0).toList();
        if (cells.isEmpty()) return null;
        return new Bounds(cells.stream().mapToInt(TablePosition::getRow).min().orElse(0), cells.stream().mapToInt(TablePosition::getRow).max().orElse(0),
                cells.stream().mapToInt(TablePosition::getColumn).min().orElse(1), cells.stream().mapToInt(TablePosition::getColumn).max().orElse(1));
    }
    private void copy() { copy(false); }
    private void copy(boolean cut) {
        if (pendingEdit != null && !pendingEdit.getAsBoolean()) return;
        var b = bounds(); if (b == null) return;
        if (!clipboardViewIsContiguous(b)) return;
        if (cut && selectedPositions().size() != (b.lastRow() - b.firstRow() + 1) * (b.lastCol() - b.firstCol() + 1)) {
            status.setText("잘라내기는 빈틈 없이 선택한 직사각형 범위에 사용할 수 있습니다."); return;
        }
        var text = new StringBuilder();
        for (int r = b.firstRow(); r <= b.lastRow(); r++) {
            if (r > b.firstRow()) text.append('\n');
            for (int c = b.firstCol(); c <= b.lastCol(); c++) {
                if (c > b.firstCol()) text.append('\t');
                String value = book.raw(sheetIndex, grid.getItems().get(r), (Integer) grid.getColumns().get(c).getUserData());
                text.append(quoteTsv(value));
            }
        }
        try {
            var snapshot = new CellClipboard(book, sheetIndex, grid.getItems().get(b.firstRow()),
                    (Integer) grid.getColumns().get(b.firstCol()).getUserData(), b.lastRow() - b.firstRow() + 1, b.lastCol() - b.firstCol() + 1);
            String id = UUID.randomUUID().toString();
            var content = new ClipboardContent(); content.putString(text.toString()); content.put(CELL_COPY, id);
            if (Clipboard.getSystemClipboard().setContent(content)) {
                cellClipboard = snapshot; copyId = id;
                pendingCut = cut ? new CutRange(sheetIndex, grid.getItems().get(b.firstRow()), (Integer) grid.getColumns().get(b.firstCol()).getUserData(), b.lastRow() - b.firstRow() + 1, b.lastCol() - b.firstCol() + 1) : null;
                if (cut) status.setText("잘라내기 대기 · 같은 시트에서 Ctrl+V로 이동 · Esc 또는 문서 편집으로 취소");
            }
        } catch (Exception e) { error(e); }
    }
    private boolean clipboardViewIsContiguous(Bounds b) {
        int firstColumn = (Integer) grid.getColumns().get(b.firstCol()).getUserData();
        for (int c = b.firstCol(); c <= b.lastCol(); c++) if ((Integer) grid.getColumns().get(c).getUserData() != firstColumn + c - b.firstCol()) {
            status.setText("열 순서를 원래대로 되돌린 뒤 복사·붙여넣기를 사용하세요."); return false;
        }
        int first = grid.getItems().get(b.firstRow());
        for (int r = b.firstRow(); r <= b.lastRow(); r++) if (grid.getItems().get(r) != first + r - b.firstRow()) {
            status.setText("복사·붙여넣기를 사용하려면 정렬·필터를 해제하고 연속 범위를 선택하세요."); return false;
        }
        return true;
    }
    private static String quoteTsv(String value) {
        return value.contains("\t") || value.contains("\n") || value.contains("\"") ? "\"" + value.replace("\"", "\"\"") + "\"" : value;
    }
    private void paste() { paste(CellClipboard.Mode.ALL); }
    private void paste(CellClipboard.Mode mode) {
        if (pendingEdit != null && !pendingEdit.getAsBoolean()) return;
        var b = bounds(); var text = Clipboard.getSystemClipboard().getString(); if (b == null || text == null) return;
        if (rangeFilter != null || !filter.getText().isBlank() || !clipboardViewIsContiguous(new Bounds(0, grid.getItems().size() - 1, b.firstCol(), b.lastCol()))) {
            status.setText("붙여넣기를 사용하려면 정렬·필터를 먼저 해제하세요."); return;
        }
        if (cellClipboard != null && copyId != null && copyId.equals(Clipboard.getSystemClipboard().getContent(CELL_COPY))) {
            if (pendingCut != null) {
                if (mode != CellClipboard.Mode.ALL) { status.setText("잘라내기는 일반 붙여넣기(Ctrl+V)를 사용하세요."); return; }
                var cut = pendingCut;
                if (cut.sheet() != sheetIndex) { status.setText("시트 사이의 잘라내기는 아직 지원하지 않습니다. 같은 시트에 붙여넣거나 복사를 사용하세요."); return; }
                pasteCut(cut, grid.getItems().get(b.firstRow()), (Integer) grid.getColumns().get(b.firstCol()).getUserData());
                return;
            }
            mutate(() -> cellClipboard.paste(book, sheetIndex, grid.getItems().get(b.firstRow()),
                    (Integer) grid.getColumns().get(b.firstCol()).getUserData(), mode)); return;
        }
        if (mode == CellClipboard.Mode.FORMATS) { status.setText("서식 붙여넣기는 이 문서에서 복사한 셀에 사용할 수 있습니다."); return; }
        try (var parser = org.apache.commons.csv.CSVFormat.TDF.parse(new StringReader(text))) {
            var rows = parser.getRecords();
            int start = b.firstRow();
            if (start + rows.size() > grid.getItems().size()) throw new IllegalArgumentException("붙여넣기가 현재 1,000행 화면을 넘습니다. 주소로 이동하거나 파일 가져오기를 사용하세요.");
            int firstCol = (Integer) grid.getColumns().get(b.firstCol()).getUserData();
            if (rows.stream().anyMatch(r -> firstCol + r.size() > Book.MAX_COLUMNS)) throw new IllegalArgumentException("최대 열 범위를 넘습니다.");
            var values = rows.stream().map(record -> record.toList()).toList();
            mutate(() -> SheetEdits.pasteText(book, sheetIndex, grid.getItems().get(start), firstCol, values,
                    mode == CellClipboard.Mode.VALUES));
        } catch (Exception e) { error(e); }
    }
    private void pasteCut(CutRange cut, int row, int column) {
        if (mutate(() -> CellMove.move(book, cut.sheet(), cut.row(), cut.column(), cut.rows(), cut.columns(), row, column))) {
            cellClipboard = null; copyId = null; status.setText("셀 이동 완료 · 참조 갱신 · Ctrl+Z로 실행 취소");
        }
    }
    private void clear() { var positions = editablePositions(); if (!positions.isEmpty()) mutate(() -> positions.forEach(p -> book.set(sheetIndex, p.row(), p.column(), ""))); }
    private void goTo() {
        try {
            var ref = new CellReference(address.getText().trim().toUpperCase(Locale.ROOT));
            int r = ref.getRow(), c = ref.getCol(); book.cell(sheetIndex, r, c, false);
            rangeFilter = null; rowWindow = (r / ROW_WINDOW) * ROW_WINDOW; columnWindow = (c / COLUMN_WINDOW) * COLUMN_WINDOW; filter.clear(); refresh();
            int index = r - rowWindow; grid.getSelectionModel().clearAndSelect(index, grid.getColumns().get(c - columnWindow + 1)); grid.scrollTo(index);
        } catch (Exception e) { error(new IllegalArgumentException("A1, Z10000처럼 올바른 셀 주소를 입력하세요.")); }
    }
    private void page(int direction) {
        if (rangeFilter != null) filterOffset = Math.clamp(filterOffset + direction * ROW_WINDOW, 0, Math.max(0, ((filterMatches.size() - 1) / ROW_WINDOW) * ROW_WINDOW));
        else rowWindow = Math.clamp(rowWindow + direction * ROW_WINDOW, 0, ((Book.MAX_ROWS - 1) / ROW_WINDOW) * ROW_WINDOW);
        refreshRows();
    }
    private void sort(boolean descending) {
        if (book.workbook().getSheetAt(sheetIndex).getNumMergedRegions() > 0) { status.setText("병합을 해제한 뒤 화면 정렬을 사용하세요."); return; }
        var p = selected(); if (p == null) return;
        Comparator<Integer> comparator = (a, b) -> {
            var x = book.raw(sheetIndex, a, p.column()); var y = book.raw(sheetIndex, b, p.column());
            if (x.isBlank() || y.isBlank()) return Boolean.compare(x.isBlank(), y.isBlank());
            int comparison;
            try { comparison = Double.compare(Double.parseDouble(x), Double.parseDouble(y)); }
            catch (NumberFormatException e) { comparison = x.compareToIgnoreCase(y); }
            return descending ? -comparison : comparison;
        };
        grid.getItems().sort(comparator); status.setText("현재 화면 정렬 · 원본 셀 주소와 수식은 유지됩니다.");
    }
    private void filterRange() {
        if (book.workbook().getSheetAt(sheetIndex).getNumMergedRegions() > 0) { status.setText("병합을 해제한 뒤 열별 필터를 사용하세요."); return; }
        if (pendingEdit != null && !pendingEdit.getAsBoolean()) return;
        var b = bounds(); if (b == null) return;
        int left = (Integer) grid.getColumns().get(b.firstCol()).getUserData();
        int right = (Integer) grid.getColumns().get(b.lastCol()).getUserData();
        var area = new TextField(rangeFilter == null ? Book.address(grid.getItems().get(b.firstRow()), left) + ":" + Book.address(grid.getItems().get(b.lastRow()), right) : rangeFilter.range().formatAsString());
        area.setId("range-filter-area");
        var header = new CheckBox("첫 행은 제목"); header.setSelected(rangeFilter == null || rangeFilter.header());
        var form = new GridPane(); form.setHgap(8); form.setVgap(10); form.addRow(0, new Label("범위 (최대 10만 행)"), area);
        form.add(header, 0, 1, 4, 1);
        var enabled = new ArrayList<CheckBox>(); var columns = new ArrayList<TextField>();
        var operations = new ArrayList<ComboBox<RangeFilter.Operation>>(); var operands = new ArrayList<TextField>();
        for (int i = 0; i < 3; i++) {
            var prior = rangeFilter != null && i < rangeFilter.conditions().size() ? rangeFilter.conditions().get(i) : null;
            var active = new CheckBox("조건 " + (i + 1)); active.setSelected(prior != null || i == 0);
            var column = new TextField(CellReference.convertNumToColString(prior == null ? left : prior.column())); column.setPrefColumnCount(3);
            var operation = new ComboBox<RangeFilter.Operation>(FXCollections.observableArrayList(RangeFilter.Operation.values()));
            operation.setValue(prior == null ? RangeFilter.Operation.CONTAINS : prior.operation());
            var operand = new TextField(prior == null ? "" : prior.operand()); operand.setPromptText("비교할 값");
            column.setId("range-filter-column-" + i); operand.setId("range-filter-value-" + i); operation.setId("range-filter-operation-" + i);
            form.addRow(i + 2, active, column, operation, operand);
            enabled.add(active); columns.add(column); operations.add(operation); operands.add(operand);
        }
        form.add(new Label("선택한 조건을 모두 만족하는 행만 표시합니다. 원본 데이터는 유지됩니다.\n문자는 대소문자 무시, 숫자 조건은 숫자 타입에만 적용합니다.\n결과는 1,000행씩 탐색하며 필터 설정은 파일에 저장하지 않습니다."), 0, 5, 4, 1);
        var dialog = new Dialog<ButtonType>(); dialog.initOwner(stage); dialog.setTitle("열별 필터");
        dialog.getDialogPane().setId("range-filter-dialog"); dialog.getDialogPane().setContent(form);
        dialog.getDialogPane().getButtonTypes().addAll(ButtonType.OK, ButtonType.CANCEL);
        dialog.showAndWait().filter(ButtonType.OK::equals).ifPresent(_ -> {
            try {
                var conditions = new ArrayList<RangeFilter.Condition>();
                for (int i = 0; i < 3; i++) if (enabled.get(i).isSelected()) conditions.add(new RangeFilter.Condition(
                        CellReference.convertColStringToIndex(columns.get(i).getText().trim().toUpperCase(Locale.ROOT)), operations.get(i).getValue(), operands.get(i).getText()));
                var spec = new RangeFilter.Spec(org.apache.poi.ss.util.CellRangeAddress.valueOf(area.getText().trim().toUpperCase(Locale.ROOT)), header.isSelected(), conditions);
                RangeFilter.matchingRows(book, sheetIndex, spec); // Validate before replacing an existing filter.
                rangeFilter = spec; filterOffset = 0; filter.clear();
                columnWindow = (spec.range().getFirstColumn() / COLUMN_WINDOW) * COLUMN_WINDOW; refresh();
            } catch (Exception e) { error(e); }
        });
    }
    private void sortRange() {
        if (pendingEdit != null && !pendingEdit.getAsBoolean()) return;
        var selected = selectedPositions(); if (selected.isEmpty()) return;
        int firstRow = selected.stream().mapToInt(Position::row).min().orElseThrow(), lastRow = selected.stream().mapToInt(Position::row).max().orElseThrow();
        int firstColumn = selected.stream().mapToInt(Position::column).min().orElseThrow(), lastColumn = selected.stream().mapToInt(Position::column).max().orElseThrow();
        var range = new TextField(Book.address(firstRow, firstColumn) + ":" + Book.address(lastRow, lastColumn));
        var column = new TextField(CellReference.convertNumToColString(firstColumn));
        range.setId("range-sort-area"); column.setId("range-sort-column");
        var header = new CheckBox("첫 행은 제목 (정렬에서 제외)"); header.setSelected(true);
        var descending = new CheckBox("내림차순");
        var form = new GridPane(); form.setHgap(10); form.setVgap(10);
        form.addRow(0, new Label("범위 (예: A1:D5000)"), range); form.addRow(1, new Label("기준 열 (예: B)"), column);
        form.add(header, 0, 2, 2, 1); form.add(descending, 0, 3, 2, 1);
        var note = new Label("지정한 직사각형의 실제 데이터를 행 단위로 재배열합니다.\n빈 값은 마지막, 같은 값은 기존 순서 유지. Ctrl+Z로 복구합니다.\n수식의 상대 참조는 새 행에 맞춰 이동하며 범위 밖 수식은 유지합니다.");
        form.add(note, 0, 4, 2, 1);
        var dialog = new Dialog<ButtonType>(); dialog.initOwner(stage); dialog.setTitle("범위 데이터 정렬");
        dialog.getDialogPane().setId("range-sort-dialog");
        dialog.getDialogPane().setContent(form); dialog.getDialogPane().getButtonTypes().addAll(ButtonType.OK, ButtonType.CANCEL);
        dialog.showAndWait().filter(ButtonType.OK::equals).ifPresent(_ -> {
            try {
                var area = org.apache.poi.ss.util.CellRangeAddress.valueOf(range.getText().trim().toUpperCase(Locale.ROOT));
                int key = CellReference.convertColStringToIndex(column.getText().trim().toUpperCase(Locale.ROOT));
                if (mutate(() -> RangeSort.sort(book, sheetIndex, area, key, header.isSelected(), descending.isSelected()))) {
                    rangeFilter = null; filter.clear(); refresh(); status.setText(area.formatAsString() + " 데이터 정렬 완료 · Ctrl+Z로 실행 취소");
                }
            } catch (Exception e) { error(e); }
        });
    }
    private void find() {
        prompt("찾기", "현재 시트에서 찾을 값 또는 수식", "").ifPresent(needle -> {
            if (needle.isEmpty()) return;
            var sheet = book.workbook().getSheetAt(sheetIndex);
            for (var row : sheet) for (var cell : row) {
                if (book.raw(sheetIndex, row.getRowNum(), cell.getColumnIndex()).toLowerCase(Locale.ROOT).contains(needle.toLowerCase(Locale.ROOT))) {
                    address.setText(cell.getAddress().formatAsString()); goTo(); return;
                }
            }
            status.setText("검색 결과가 없습니다.");
        });
    }
    private void fill(SheetEdits.Direction direction) {
        if (rangeFilter != null) { status.setText("채우기를 사용하려면 열별 필터를 해제하세요."); return; }
        var b = bounds(); if (b == null) return;
        int first = grid.getItems().get(b.firstRow()), last = grid.getItems().get(b.lastRow());
        for (int r = b.firstRow(); r <= b.lastRow(); r++) {
            if (grid.getItems().get(r) != first + r - b.firstRow()) { status.setText("채우기를 사용하려면 정렬·필터를 해제하고 연속된 범위를 선택하세요."); return; }
        }
        int left = (Integer) grid.getColumns().get(b.firstCol()).getUserData(), right = (Integer) grid.getColumns().get(b.lastCol()).getUserData();
        mutate(() -> SheetEdits.fill(book, sheetIndex, first, last, left, right, direction));
    }
    private void structuralEdit(StructuralEdits.Axis axis, boolean insert) {
        if (pendingEdit != null && !pendingEdit.getAsBoolean()) return;
        var positions = selectedPositions(); if (positions.isEmpty()) return;
        if (rangeFilter != null || !filter.getText().isBlank() || !grid.getSortOrder().isEmpty()) {
            status.setText("행·열 삽입·삭제를 사용하려면 정렬과 필터를 먼저 해제하세요."); return;
        }
        var indexes = positions.stream().mapToInt(p -> axis == StructuralEdits.Axis.ROW ? p.row() : p.column()).distinct().sorted().toArray();
        int first = indexes[0], count = indexes.length;
        if (indexes[count - 1] - first + 1 != count) { status.setText("연속된 행 또는 열을 선택하세요."); return; }
        if (mutate(() -> StructuralEdits.change(book, sheetIndex, axis, first, count, insert))) {
            refresh();
            status.setText(count + (axis == StructuralEdits.Axis.ROW ? "개 행" : "개 열") + (insert ? " 삽입 완료" : " 삭제 완료") + " · Ctrl+Z로 실행 취소");
        }
    }
    private void replaceAll() {
        var dialog = new Dialog<ButtonType>(); dialog.initOwner(stage); dialog.setTitle("찾기 및 바꾸기");
        var find = new TextField(); find.setPromptText("찾을 텍스트"); var replacement = new TextField(); replacement.setPromptText("바꿀 텍스트");
        var matchCase = new CheckBox("대소문자 구분"); var formulas = new CheckBox("수식도 검색 (참조가 바뀔 수 있음)");
        var form = new VBox(10, new Label("현재 시트의 텍스트를 모두 바꿉니다. 실행 취소할 수 있습니다."), find, replacement, matchCase, formulas); form.setPadding(new Insets(12));
        var replace = new ButtonType("모두 바꾸기", ButtonBar.ButtonData.OK_DONE); dialog.getDialogPane().getButtonTypes().addAll(replace, ButtonType.CANCEL); dialog.getDialogPane().setContent(form);
        if (dialog.showAndWait().orElse(ButtonType.CANCEL) == replace) {
            mutate(() -> { int count = SheetEdits.replaceAll(book, sheetIndex, find.getText(), replacement.getText(), matchCase.isSelected(), formulas.isSelected()); status.setText(count + "개 셀을 바꿨습니다."); });
        }
    }
    private void rowHeight() {
        if (pendingEdit != null && !pendingEdit.getAsBoolean()) return;
        prompt("행 높이", "높이 (1–409 포인트)", "30").ifPresent(value -> {
            try { applyRowHeight(Double.parseDouble(value)); } catch (Exception e) { error(e); }
        });
    }
    private void applyRowHeight(Double points) {
        var rows = selectedPositions().stream().map(Position::row).distinct().toList();
        if (!rows.isEmpty()) mutate(() -> rows.forEach(r -> book.setRowHeight(sheetIndex, r, points)));
    }
    private void columnWidth() {
        var positions = selectedPositions(); if (positions.isEmpty()) return;
        prompt("열 너비", "문자 기준 너비 (1–255)", "15").ifPresent(value -> {
            try {
                int width = Integer.parseInt(value); if (width < 1 || width > 255) throw new IllegalArgumentException("1–255 범위의 너비를 입력하세요.");
                mutate(() -> positions.stream().map(Position::column).distinct().forEach(c -> book.workbook().getSheetAt(sheetIndex).setColumnWidth(c, width * 256))); refresh();
            } catch (Exception e) { error(e); }
        });
    }
    private void addSheet() {
        prompt("시트 추가", "새 시트 이름", "Sheet" + (book.workbook().getNumberOfSheets() + 1)).ifPresent(name -> {
            mutate(() -> { WorkbookUtil.validateSheetName(name); book.workbook().createSheet(name); });
            rangeFilter = null; sheetIndex = book.workbook().getNumberOfSheets() - 1; refreshTabs(); refresh();
        });
    }
    private void renameSheet() {
        prompt("시트 이름 변경", "새 이름", book.workbook().getSheetName(sheetIndex)).ifPresent(name -> {
            mutate(() -> { WorkbookUtil.validateSheetName(name); book.workbook().setSheetName(sheetIndex, name); }); refreshTabs();
        });
    }
    private void deleteSheet() {
        if (book.workbook().getNumberOfSheets() == 1) { status.setText("시트는 최소 한 개 필요합니다."); return; }
        if (confirm("시트 삭제", "현재 시트를 삭제할까요? 실행 취소로 복구할 수 있습니다.")) {
            rangeFilter = null; int deleting = sheetIndex; sheetIndex = 0;
            mutate(() -> book.workbook().removeSheetAt(deleting)); refreshTabs(); refresh();
        }
    }
    private void newBook() {
        if (!canLeave()) return;
        replace(new Book()); document = null; dirty = false; sheetIndex = rowWindow = columnWindow = 0; filter.clear(); refreshTabs(); refresh();
    }
    private void open() {
        if (!canLeave()) return;
        var chooser = chooser("문서 열기", "스프레드시트", "*.gsheet", "*.xlsx", "*.csv");
        var file = chooser.showOpenDialog(stage); if (file == null) return;
        try {
            var path = file.toPath(); String name = file.getName().toLowerCase(Locale.ROOT);
            if (name.endsWith(".xlsx") && !confirm("Excel 가져오기", ".gsheet로 저장할 때 고급 Excel 기능(그림, 차트, 이름 정의, 조건부 서식 등)이 제외됩니다. 원본은 수정하지 않습니다. 계속할까요?")) return;
            var loaded = name.endsWith(".xlsx") ? BookFiles.importExcel(path) : name.endsWith(".csv") ? BookFiles.importCsv(path) : BookFiles.read(path);
            replace(loaded); document = name.endsWith(".gsheet") ? path : null; dirty = document == null;
            sheetIndex = rowWindow = columnWindow = 0; filter.clear(); refreshTabs(); refresh();
        } catch (Exception e) { error(e); }
    }
    private boolean save(boolean saveAs) {
        if (gitBusy) { status.setText("Git 작업을 마친 후 다시 저장하세요."); return false; }
        if (pendingEdit != null && !pendingEdit.getAsBoolean()) return false;
        Path target = document;
        if (saveAs || target == null) {
            var chooser = chooser("Git Sheet 저장", "Git Sheet", "*.gsheet"); chooser.setInitialFileName(document == null ? "workbook.gsheet" : document.getFileName().toString());
            var file = chooser.showSaveDialog(stage); if (file == null) return false; target = extension(file.toPath(), ".gsheet");
        }
        try { BookFiles.write(book, target); document = target.toAbsolutePath().normalize(); dirty = false; title(); status.setText("저장됨 · " + document); saveNotification.show(document); return true; }
        catch (Exception e) { error(e); return false; }
    }
    private void export(boolean csv) {
        var chooser = chooser("내보내기", csv ? "CSV" : "Excel", csv ? "*.csv" : "*.xlsx");
        chooser.setInitialFileName(csv ? "sheet.csv" : "workbook.xlsx"); var file = chooser.showSaveDialog(stage); if (file == null) return;
        try {
            var path = extension(file.toPath(), csv ? ".csv" : ".xlsx");
            if (csv) BookFiles.exportCsv(book, sheetIndex, path); else BookFiles.exportExcel(book, path);
            status.setText("내보내기 완료 · " + path);
            saveNotification.show(path);
        } catch (Exception e) { error(e); }
    }
    private static Path extension(Path path, String extension) { return path.toString().toLowerCase(Locale.ROOT).endsWith(extension) ? path : path.resolveSibling(path.getFileName() + extension); }
    private static FileChooser chooser(String title, String label, String... patterns) {
        var chooser = new FileChooser(); chooser.setTitle(title); chooser.getExtensionFilters().add(new FileChooser.ExtensionFilter(label, patterns)); return chooser;
    }
    private boolean canLeave() {
        if (pendingEdit != null && !pendingEdit.getAsBoolean()) return false;
        if (!dirty) return true;
        var alert = new Alert(Alert.AlertType.CONFIRMATION, "저장하지 않은 변경이 있습니다.", ButtonType.YES, ButtonType.NO, ButtonType.CANCEL);
        alert.initOwner(stage); alert.setHeaderText("변경 내용을 저장할까요?");
        var answer = alert.showAndWait().orElse(ButtonType.CANCEL);
        return answer == ButtonType.YES ? save(false) : answer == ButtonType.NO;
    }
    private void replace(Book next) {
        rangeFilter = null;
        pendingCut = null;
        cellClipboard = null; copyId = null;
        try { book.close(); } catch (IOException ignored) { }
        book = next;
    }
    private void gitInspect() {
        if (gitBusy) return;
        if (document == null || dirty) { if (!save(false)) return; }
        Path path = document; async("Git 변경 내역", () -> { try { return GitService.inspect(path); } catch (Exception e) { throw new CompletionException(e); } });
    }
    private void gitCommit() {
        if (gitBusy) return;
        prompt("Git 커밋", "선택한 문서만 저장하고 커밋합니다. 새 저장소는 문서 폴더에 생성됩니다.", "Update workbook").ifPresent(message -> {
            if (!save(false)) return;
            Path path = document; async("Git 커밋", () -> { try { return GitService.commit(path, message); } catch (Exception e) { throw new CompletionException(e); } });
        });
    }
    private void async(String title, Supplier<String> task) {
        if (gitBusy) return;
        gitBusy = true;
        status.setText(title + " 실행 중…");
        CompletableFuture.supplyAsync(task, background).whenComplete((result, error) -> Platform.runLater(() -> {
            gitBusy = false;
            status.setText(title + (error == null ? " 완료" : " 실패"));
            if (error != null) error(error.getCause() == null ? error : error.getCause()); else textDialog(title, result);
        }));
    }
    private void gitConnect() {
        if (document == null && !save(false)) return;
        if (gitDialog != null && gitDialog.isShowing()) { gitDialog.getDialogPane().getScene().getWindow().requestFocus(); return; }
        gitDialog = new GitDialog(stage, document, background); gitDialog.show();
    }
    private void chart() {
        var b = bounds(); if (b == null || b.firstCol() == b.lastCol()) { status.setText("차트를 만들려면 이름 열과 숫자 열을 함께 선택하세요."); return; }
        var chart = new BarChart<String, Number>(new CategoryAxis(), new NumberAxis()); chart.setTitle("선택 영역 차트"); chart.setLegendVisible(false); chart.setAnimated(false);
        var series = new XYChart.Series<String, Number>();
        for (int r = b.firstRow(); r <= Math.min(b.lastRow(), b.firstRow() + 199); r++) {
            int row = grid.getItems().get(r), labelCol = (Integer) grid.getColumns().get(b.firstCol()).getUserData(), valueCol = (Integer) grid.getColumns().get(b.lastCol()).getUserData();
            try {
                var cell = book.cell(sheetIndex, row, valueCol, false); if (cell == null) continue;
                var evaluator = book.workbook().getCreationHelper().createFormulaEvaluator();
                var evaluated = evaluator.evaluate(cell); if (evaluated.getCellType() != CellType.NUMERIC) continue;
                series.getData().add(new XYChart.Data<>(book.display(sheetIndex, row, labelCol) + " (" + (row + 1) + ")", evaluated.getNumberValue()));
            } catch (RuntimeException ignored) { }
        }
        chart.getData().add(series); var window = new Stage(); window.initOwner(stage); window.setTitle("차트 · 최대 200행 · 저장되지 않는 미리보기"); window.setScene(new Scene(chart, 850, 520)); window.show();
    }
    private Optional<String> prompt(String title, String description, String initial) {
        var dialog = new TextInputDialog(initial); dialog.initOwner(stage); dialog.setTitle(title); dialog.setHeaderText(description); return dialog.showAndWait();
    }
    private boolean confirm(String title, String message) {
        var dialog = new Alert(Alert.AlertType.CONFIRMATION, message, ButtonType.OK, ButtonType.CANCEL); dialog.initOwner(stage); dialog.setTitle(title); dialog.setHeaderText(null); return dialog.showAndWait().orElse(ButtonType.CANCEL) == ButtonType.OK;
    }
    private void textDialog(String title, String text) {
        var dialog = new Dialog<Void>(); dialog.initOwner(stage); dialog.setTitle(title); dialog.getDialogPane().getButtonTypes().add(ButtonType.CLOSE);
        var area = new TextArea(text); area.setEditable(false); area.setWrapText(true); area.setPrefSize(780, 480); dialog.getDialogPane().setContent(area); dialog.setResizable(true); dialog.show();
    }
    private void error(Throwable error) { var a = new Alert(Alert.AlertType.ERROR, Objects.toString(error.getMessage(), error.getClass().getSimpleName()), ButtonType.OK); a.initOwner(stage); a.setHeaderText("작업을 완료하지 못했습니다."); a.showAndWait(); }
    private void title() { stage.setTitle((dirty ? "● " : "") + (document == null ? "새 문서" : document.getFileName()) + " — Git Sheet"); }
    private void sample() {
        String[][] data = {{"품목", "수량", "단가", "금액"}, {"키보드", "3", "120000", "=B2*C2"}, {"마우스", "4", "50000", "=B3*C3"}, {"허브", "2", "50000", "=B4*C4"}, {"합계", "=SUM(B2:B4)", "", "=SUM(D2:D4)"}};
        for (int r = 0; r < data.length; r++) for (int c = 0; c < data[r].length; c++) {
            book.set(0, r, c, data[r][c]);
            if (r == 0 || r == 4) { book.style(0, r, c, "bold", "true"); book.style(0, r, c, "fill", Short.toString(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex())); }
            if (r > 0 && c >= 2) book.style(0, r, c, "format", "₩#,##0");
        }
        dirty = true;
    }
    /** Integration smoke exercises actual table edit controls, history, filter and navigation. */
    private void smokeTest() {
        grid.getSelectionModel().clearAndSelect(1, grid.getColumns().get(2));
        grid.getFocusModel().focus(1, grid.getColumns().get(2)); grid.requestFocus();
        grid.fireEvent(new KeyEvent(KeyEvent.KEY_TYPED, "5", "", KeyCode.UNDEFINED, false, false, false, false));
        grid.applyCss(); grid.layout();
        var editor = grid.lookupAll(".text-area").stream().filter(n -> n instanceof TextArea && n.isVisible()).map(n -> (TextArea)n).findFirst().orElseThrow();
        if (!editor.getText().equals("5")) throw new IllegalStateException("Type-to-replace failed");
        editor.fireEvent(new KeyEvent(KeyEvent.KEY_PRESSED, "", "", KeyCode.ENTER, false, false, false, false));
        if (!book.raw(0, 1, 1).equals("5") || !book.display(0, 1, 3).equals("₩600,000")) throw new IllegalStateException("Edit/recalculation failed");
        undo(); if (!book.raw(0, 1, 1).equals("3")) throw new IllegalStateException("Undo failed");
        redo(); if (!book.raw(0, 1, 1).equals("5")) throw new IllegalStateException("Redo failed");
        undo(); smokeLayout();
        grid.getSelectionModel().clearAndSelect(1, grid.getColumns().get(2));
        grid.getFocusModel().focus(1, grid.getColumns().get(2));
        grid.fireEvent(new KeyEvent(KeyEvent.KEY_TYPED, "9", "", KeyCode.UNDEFINED, false, false, false, false));
        var typedEditor = (TextArea) smokeCell(1, 1).getGraphic();
        typedEditor.fireEvent(new KeyEvent(KeyEvent.KEY_PRESSED, "", "", KeyCode.ESCAPE, false, false, false, false));
        if (!book.raw(0, 1, 1).equals("3")) throw new IllegalStateException("Type/Escape changed value");
        grid.fireEvent(new KeyEvent(KeyEvent.KEY_PRESSED, "", "", KeyCode.F2, false, false, false, false));
        typedEditor = (TextArea) smokeCell(1, 1).getGraphic();
        if (!typedEditor.getText().equals("3")) throw new IllegalStateException("F2 lost existing value");
        typedEditor.fireEvent(new KeyEvent(KeyEvent.KEY_PRESSED, "", "", KeyCode.ESCAPE, false, false, false, false));
        grid.fireEvent(new InputMethodEvent(InputMethodEvent.INPUT_METHOD_TEXT_CHANGED, List.of(), "한글", 0));
        typedEditor = (TextArea) smokeCell(1, 1).getGraphic();
        if (!typedEditor.getText().equals("한글")) throw new IllegalStateException("IME committed text lost");
        typedEditor.fireEvent(new KeyEvent(KeyEvent.KEY_PRESSED, "", "", KeyCode.ESCAPE, false, false, false, false));
        // Avoid changing the user's system clipboard during smoke checks.
        filter.setText("마우스"); refreshRows();
        if (grid.getItems().size() != 1 || grid.getItems().getFirst() != 2) throw new IllegalStateException("Filter failed");
        filter.clear(); address.setText("AZ2001"); goTo();
        if (rowWindow != 2000 || !Objects.equals(selected(), new Position(2000, 51))) throw new IllegalStateException("Navigation failed");
        address.setText("A1"); goTo(); grid.applyCss(); grid.layout();
        Platform.runLater(() -> {
            try {
                var pane = Window.getWindows().stream().filter(w -> w.getScene() != null)
                        .map(w -> w.getScene().lookup("#range-sort-dialog")).filter(n -> n instanceof DialogPane)
                        .map(n -> (DialogPane) n).findFirst().orElseThrow();
                ((TextField) pane.lookup("#range-sort-area")).setText("A1:D4");
                ((TextField) pane.lookup("#range-sort-column")).setText("B");
                ((Button) pane.lookupButton(ButtonType.OK)).fire();
            } catch (Exception e) { failSmoke(e); }
        });
        sortRange();
        if (!book.raw(0, 1, 0).equals("허브") || !book.raw(0, 1, 3).equals("=B2*C2") || !book.raw(0, 0, 0).equals("품목"))
            throw new IllegalStateException("Range sort dialog failed");
        undo(); address.setText("A1"); goTo();
        String original = book.raw(0, 0, 0);
        book.transaction(x -> { x.set(0, 0, 0, "줄바꿈 text ".repeat(30)); x.style(0, 0, 0, "wrap", "true"); });
        refresh(); smokeLayout();
        double narrow = smokeCell(0, 0).getTableRow().getHeight();
        grid.getColumns().get(1).setPrefWidth(400); grid.refresh(); smokeLayout();
        double wide = smokeCell(0, 0).getTableRow().getHeight();
        if (narrow <= wide || wide < 40) throw new IllegalStateException("Wrap height did not follow column width: " + narrow + "/" + wide);
        grid.getSelectionModel().clearAndSelect(0, grid.getColumns().get(1)); applyRowHeight(30.0); smokeLayout();
        if (Math.abs(smokeCell(0, 0).getTableRow().getHeight() - 40) > 2) throw new IllegalStateException("Manual row height failed");
        applyRowHeight(null); smokeLayout();
        if (smokeCell(0, 0).getTableRow().getHeight() < 50) throw new IllegalStateException("Automatic row height failed");
        undo(); undo(); undo(); address.setText("A1"); goTo();
        grid.edit(0, grid.getColumns().get(1)); smokeLayout();
        var multiline = smokeCell(0, 0).editor; multiline.setText("line 1"); multiline.positionCaret(multiline.getLength());
        multiline.fireEvent(new KeyEvent(KeyEvent.KEY_PRESSED, "", "", KeyCode.ENTER, false, false, true, false));
        multiline.appendText("line 2");
        if (!pendingEdit.getAsBoolean() || !book.raw(0, 0, 0).equals("line 1\nline 2")) throw new IllegalStateException("Multiline editor failed");
        formula.fireEvent(new KeyEvent(KeyEvent.KEY_PRESSED, "", "", KeyCode.ENTER, false, false, false, false));
        if (!book.raw(0, 0, 0).equals("line 1\nline 2")) throw new IllegalStateException("Formula bar lost line breaks");
        undo(); undo(); address.setText("A1"); goTo();
        Platform.runLater(() -> {
            try {
                var pane = Window.getWindows().stream().filter(w -> w.getScene() != null)
                        .map(w -> w.getScene().lookup("#range-filter-dialog")).filter(n -> n instanceof DialogPane)
                        .map(n -> (DialogPane) n).findFirst().orElseThrow();
                ((TextField) pane.lookup("#range-filter-area")).setText("A1:D4");
                ((TextField) pane.lookup("#range-filter-column-0")).setText("A");
                ((TextField) pane.lookup("#range-filter-value-0")).setText("마우스");
                ((Button) pane.lookupButton(ButtonType.OK)).fire();
            } catch (Exception e) { failSmoke(e); }
        });
        filterRange();
        if (!grid.getItems().equals(List.of(0, 2))) throw new IllegalStateException("Column filter dialog failed");
        mutate(() -> book.set(0, 2, 0, "renamed"));
        if (!grid.getItems().equals(List.of(0))) throw new IllegalStateException("Filter did not refresh after editing");
        undo(); if (!grid.getItems().equals(List.of(0, 2))) throw new IllegalStateException("Filter undo refresh failed");
        book.transaction(x -> x.workbook().createSheet("Filter smoke")); refreshTabs(); tabs.getSelectionModel().select(1);
        if (rangeFilter != null) throw new IllegalStateException("Filter survived sheet switch");
        undo(); address.setText("A1"); goTo();
        rangeFilter = new RangeFilter.Spec(org.apache.poi.ss.util.CellRangeAddress.valueOf("A1:A2005"), true,
                List.of(new RangeFilter.Condition(0, RangeFilter.Operation.BLANK, "")));
        filterOffset = 0; refreshRows();
        if (grid.getItems().get(1) != 5) throw new IllegalStateException("Filter first result page failed");
        page(1); if (grid.getItems().get(1) != 1005) throw new IllegalStateException("Filter next result page failed");
        rangeFilter = null; refreshRows();
        if (grid.getItems().get(1) != 1 || filter.isDisabled()) throw new IllegalStateException("Filter clear failed");
        pendingCut = new CutRange(0, 0, 0, 1, 1);
        pasteCut(pendingCut, 15, 0);
        if (pendingCut != null || !book.raw(0, 0, 0).isEmpty() || !book.raw(0, 15, 0).equals(original)) throw new IllegalStateException("Cut paste failed");
        undo(); address.setText("A1"); goTo();
        structuralEdit(StructuralEdits.Axis.ROW, true);
        if (!book.raw(0, 1, 0).equals(original) || !book.raw(0, 0, 0).isEmpty()) throw new IllegalStateException("Insert row failed");
        undo(); address.setText("A1"); goTo();
        book.transaction(x -> x.set(0, 0, 0, "'=literal")); refresh();
        grid.getSelectionModel().clearAndSelect(0, grid.getColumns().get(1));
        grid.edit(0, grid.getColumns().get(1)); grid.applyCss(); grid.layout();
        if (pendingEdit == null || !pendingEdit.getAsBoolean() || book.cell(0, 0, 0, false).getCellType() != CellType.STRING
                || !book.raw(0, 0, 0).equals("=literal")) throw new IllegalStateException("Literal text editing failed");
        undo(); undo(); address.setText("A1"); goTo();
        structuralEdit(StructuralEdits.Axis.COLUMN, true);
        if (!book.raw(0, 0, 1).equals(original)) throw new IllegalStateException("Insert column failed");
        undo(); address.setText("A1"); goTo();
    }
    private String cellCss(org.apache.poi.ss.usermodel.Cell cell, boolean selected) {
        if (cell == null) return "";
            var style = cell.getCellStyle(); var font = book.workbook().getFontAt(style.getFontIndex());
            var css = new StringBuilder("-fx-font-weight:").append(font.getBold() ? "bold" : "normal").append(";-fx-font-style:").append(font.getItalic() ? "italic" : "normal").append(';');
            if (style instanceof org.apache.poi.xssf.usermodel.XSSFCellStyle xs && style.getFillPattern() == FillPatternType.SOLID_FOREGROUND && !selected) {
                var color = xs.getFillForegroundXSSFColor(); var rgb = color == null ? null : color.getRGB();
                if (rgb != null) css.append("-fx-background-color:#").append(HexFormat.of().formatHex(rgb)).append(';');
            }
            css.append("-fx-font-size:").append(font.getFontHeightInPoints()).append("pt;");
            String alignment = switch (style.getAlignment()) {
                case CENTER, CENTER_SELECTION -> "CENTER"; case RIGHT -> "CENTER-RIGHT"; case LEFT -> "CENTER-LEFT";
                default -> cell.getCellType() == CellType.NUMERIC || cell.getCellType() == CellType.FORMULA ? "CENTER-RIGHT" : "CENTER-LEFT";
            };
            css.append("-fx-alignment:").append(alignment).append(';');
            if (style.getBorderBottom() != BorderStyle.NONE || style.getBorderTop() != BorderStyle.NONE || style.getBorderLeft() != BorderStyle.NONE || style.getBorderRight() != BorderStyle.NONE)
                css.append("-fx-border-color:#64748b;-fx-border-width:1;");
            return css.toString();
    }
    private void smokeLayout() { stage.getScene().getRoot().applyCss(); stage.getScene().getRoot().layout(); grid.layout(); paintMerges(); }
    private GridCell smokeCell(int row, int column) {
        return grid.lookupAll(".table-cell").stream().filter(n -> n instanceof GridCell)
                .map(n -> (GridCell) n).filter(c -> c.getTableRow().getItem() != null && c.getTableRow().getItem() == row && c.column == column)
                .findFirst().orElseThrow();
    }
    private void screenshot(Path path) throws IOException {
        var image = stage.getScene().getRoot().snapshot(null, null);
        var rendered = new java.awt.image.BufferedImage((int) image.getWidth(), (int) image.getHeight(), java.awt.image.BufferedImage.TYPE_INT_ARGB);
        var pixels = image.getPixelReader();
        for (int y = 0; y < rendered.getHeight(); y++) for (int x = 0; x < rendered.getWidth(); x++) rendered.setRGB(x, y, pixels.getArgb(x, y));
        if (path.toAbsolutePath().getParent() != null) Files.createDirectories(path.toAbsolutePath().getParent());
        javax.imageio.ImageIO.write(rendered, "png", path.toFile());
    }
    @SuppressWarnings("unchecked")
    private void smokeMergedCells() throws IOException {
        address.setText("A1"); goTo(); smokeLayout();
        book.transaction(b -> { b.set(0, 9, 0, "Merged title"); b.style(0, 9, 0, "bold", "true"); });
        grid.refresh(); smokeLayout();
        grid.getSelectionModel().clearSelection();
        grid.getSelectionModel().selectRange(9, grid.getColumns().get(1), 10, grid.getColumns().get(2));
        actions.get("mergeCells").run(); smokeLayout();
        if (mergeLayer.getChildren().size() != 1) throw new IllegalStateException("Merge overlay missing");
        var label = (Label) mergeLayer.getChildren().getFirst();
        double expectedWidth = smokeCell(9, 0).getWidth() + smokeCell(9, 1).getWidth();
        double expectedHeight = smokeCell(9, 0).getHeight() + smokeCell(10, 0).getHeight();
        if (!label.getText().equals("Merged title") || Math.abs(label.getWidth() - expectedWidth) > 2
                || Math.abs(label.getHeight() - expectedHeight) > 2)
            throw new IllegalStateException("Merge span incorrect: " + label.getWidth() + "x" + label.getHeight()
                    + " expected " + expectedWidth + "x" + expectedHeight);
        double originalColumnWidth = grid.getColumns().get(1).getPrefWidth();
        double originalMergeWidth = label.getWidth();
        grid.getColumns().get(1).setPrefWidth(originalColumnWidth + 80); smokeLayout();
        if (Math.abs(label.getWidth() - originalMergeWidth - 80) > 2) throw new IllegalStateException("Merge did not resize with column");
        grid.getColumns().get(1).setPrefWidth(originalColumnWidth); smokeLayout();
        screenshot(Path.of("build/merged-cell-smoke.png"));
        grid.scrollTo(50); smokeLayout();
        if (!mergeLayer.getChildren().isEmpty()) throw new IllegalStateException("Offscreen merge remained visible");
        grid.scrollTo(9); smokeLayout();
        if (mergeLayer.getChildren().isEmpty()) throw new IllegalStateException("Scrolled merge did not return");
        grid.getSelectionModel().clearAndSelect(10, grid.getColumns().get(2));
        grid.getFocusModel().focus(10, grid.getColumns().get(2));
        grid.fireEvent(new KeyEvent(KeyEvent.KEY_TYPED, "병합", "", KeyCode.UNDEFINED, false, false, false, false));
        if (pendingEdit == null || !pendingEdit.getAsBoolean()) throw new IllegalStateException("Merged typing failed");
        if (!book.raw(0, 9, 0).equals("병합") || !book.raw(0, 10, 1).isEmpty()) throw new IllegalStateException("Merged edit missed anchor");
        undo(); smokeLayout();
        grid.getSelectionModel().clearAndSelect(10, grid.getColumns().get(2));
        actions.get("unmergeCells").run(); smokeLayout();
        if (!mergeLayer.getChildren().isEmpty() || book.workbook().getSheetAt(0).getNumMergedRegions() != 0)
            throw new IllegalStateException("Unmerge failed");
        undo(); undo(); undo(); address.setText("A1"); goTo(); smokeLayout();
        book.transaction(b -> {
            b.set(0, 998, 0, "Across pages");
            MergeEdits.merge(b, 0, org.apache.poi.ss.util.CellRangeAddress.valueOf("A999:B1001"));
        });
        address.setText("B1001"); goTo(); smokeLayout();
        if (rowWindow != 1000 || mergeLayer.getChildren().size() != 1
                || !((Label) mergeLayer.getChildren().getFirst()).getText().equals("Across pages"))
            throw new IllegalStateException("Merge fragment at page boundary missing");
        grid.getFocusModel().focus(0, grid.getColumns().get(2));
        grid.fireEvent(new KeyEvent(KeyEvent.KEY_TYPED, "boundary", "", KeyCode.UNDEFINED, false, false, false, false));
        if (pendingEdit == null || !pendingEdit.getAsBoolean() || rowWindow != 0
                || !book.raw(0, 998, 0).equals("boundary") || !book.raw(0, 1000, 1).isEmpty())
            throw new IllegalStateException("Page-boundary editing did not reach merge anchor");
        undo(); undo(); address.setText("A1"); goTo(); smokeLayout();
    }
    private void smokeSaveAndGit() throws Exception {
        var folder = Files.createTempDirectory(Path.of("build"), "ui-smoke-").toAbsolutePath();
        if (GitService.run(folder, "init", "-b", "main").code() != 0) throw new IllegalStateException("Smoke repository initialization failed");
        document = folder.resolve("저장 알림.gsheet");
        if (!save(false) || !saveNotification.isVisible() || !saveNotification.displayedPath().equals(document.toString())) throw new IllegalStateException("Save notification did not display the actual path");
        gitDialog = new GitDialog(stage, document, background); gitDialog.show();
        gitDialog.whenIdle().thenRun(() -> {
            var pane = gitDialog.getDialogPane();
            var services = (ComboBox<GitService.Provider>) pane.lookup("#git-provider");
            services.setValue(GitService.Provider.GITEA);
            ((TextField) pane.lookup("#git-url")).setText("https://gitea.example/test/workbook.git");
            ((TextField) pane.lookup("#git-name")).setText("UI Smoke");
            ((TextField) pane.lookup("#git-email")).setText("smoke@example.invalid");
            ((Button) pane.lookup("#git-save")).fire();
            gitDialog.whenIdle().thenRun(() -> {
                try {
                    var settings = GitService.settings(document);
                    if (settings.provider() != GitService.Provider.GITEA || !settings.url().equals("https://gitea.example/test/workbook.git")) throw new IllegalStateException("Provider settings did not persist");
                    gitDialog.close(); save(false);
                    var image = getParameters().getNamed().get("screenshot");
                    stage.getScene().getRoot().applyCss(); stage.getScene().getRoot().layout();
                    if (image != null) screenshot(Path.of(image));
                    var delay = new javafx.animation.PauseTransition(SaveNotification.DISPLAY_TIME.add(javafx.util.Duration.millis(500)));
                    delay.setOnFinished(_ -> {
                        if (saveNotification.isVisible()) { failSmoke(new IllegalStateException("Save notification did not auto-dismiss")); return; }
                        System.out.println("GIT_SHEET_UI_SMOKE_OK (editing, save popup, provider persistence, timed dismissal)"); Platform.exit();
                    }); delay.play();
                } catch (Exception e) { failSmoke(e); }
            }).exceptionally(failure -> { failSmoke(failure); return null; });
        }).exceptionally(failure -> { failSmoke(failure); return null; });
    }
    private static void failSmoke(Throwable failure) { failure.printStackTrace(); System.exit(1); }
    private final class GridCell extends TableCell<Integer, String> {
        private final int column;
        private TextArea editor;
        GridCell(int column) { this.column = column; }
        @Override protected void updateItem(String value, boolean empty) {
            super.updateItem(value, empty); requestMergePaint();
            setWrapText(false);
            if (empty) { setText(null); setGraphic(null); setStyle(""); return; }
            if (!isEditing()) { setText(value); setGraphic(null); }
            int row = getTableRow().getItem() == null ? -1 : getTableRow().getItem();
            var cell = row < 0 ? null : book.cell(sheetIndex, row, column, false);
            if (cell == null) { setStyle(""); return; }
            setWrapText(cell.getCellStyle().getWrapText());
            setStyle(cellCss(cell, isSelected()));
        }
        @Override public void startEdit() {
            if (isEmpty()) return;
            int row = getTableRow().getItem();
            var anchor = mergeAnchor(row, column);
            if (anchor.row() != row || anchor.column() != column) { focusMergeAnchor(anchor, true); return; }
            super.startEdit();
            editor = new TextArea(book.input(sheetIndex, row, column)); editor.setWrapText(true); editor.setPrefRowCount(2);
            pendingEdit = () -> finish(row);
            editor.addEventFilter(KeyEvent.KEY_PRESSED, event -> {
                if (event.getCode() == KeyCode.ESCAPE) { cancelEdit(); event.consume(); }
                else if (event.getCode() == KeyCode.ENTER) {
                    if (event.isAltDown()) editor.replaceSelection("\n"); else finish(row);
                    event.consume();
                }
            });
            editor.focusedProperty().addListener((_, _, focus) -> { if (!focus && isEditing()) finish(row); });
            setText(null); setGraphic(editor); editor.requestFocus(); editor.selectAll();
        }
        private boolean finish(int row) {
            if (!isEditing()) { pendingEdit = null; return true; }
            String text = editor.getText(); pendingEdit = null; super.commitEdit(text); setGraphic(null);
            return mutate(() -> book.set(sheetIndex, row, column, text));
        }
        @Override public void cancelEdit() { pendingEdit = null; super.cancelEdit(); setGraphic(null); setText(getItem()); }
        @Override protected double computePrefHeight(double width) {
            double available = getWidth() > 0 ? getWidth() : getTableColumn() == null ? width : getTableColumn().getWidth();
            return Math.min(409 * 96.0 / 72, super.computePrefHeight(isWrapText() ? Math.max(24, available) : width));
        }
    }
    @Override public void stop() throws Exception { if (mergeOverlay != null) mergeOverlay.close(); background.shutdownNow(); book.close(); }
}
