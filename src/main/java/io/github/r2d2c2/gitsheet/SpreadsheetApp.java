package io.github.r2d2c2.gitsheet;

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
    private Stage stage;
    private Path document;
    private boolean dirty;
    private int sheetIndex, rowWindow, columnWindow;
    private static final int ROW_WINDOW = 1000, COLUMN_WINDOW = 52;
    private final TableView<Integer> grid = new TableView<>();
    private final TabPane tabs = new TabPane();
    private final TextField address = new TextField("A1"), formula = new TextField(), filter = new TextField();
    private final Label status = new Label("준비"), selectionInfo = new Label();
    private final ExecutorService background = Executors.newVirtualThreadPerTaskExecutor();
    private boolean changingTabs;
    private SaveNotification saveNotification;
    private GitDialog gitDialog;
    private boolean gitBusy;
    private BooleanSupplier pendingEdit;
    private record Position(int row, int column) {}
    private record Bounds(int firstRow, int lastRow, int firstCol, int lastCol) {}

    public static void main(String[] args) { launch(args); }
    @Override public void start(Stage stage) {
        this.stage = stage;
        var root = new BorderPane();
        root.setTop(new VBox(menu(), toolbar(), formulaBar()));
        grid.setEditable(true); grid.setFixedCellSize(-1); grid.getSelectionModel().setCellSelectionEnabled(true);
        grid.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        grid.getSelectionModel().getSelectedCells().addListener((javafx.collections.ListChangeListener<TablePosition>) change -> selectionChanged());
        grid.setOnKeyPressed(this::gridKey);
        grid.setPlaceholder(new Label("조건에 맞는 행이 없습니다."));
        root.setCenter(grid);
        tabs.setTabClosingPolicy(TabPane.TabClosingPolicy.UNAVAILABLE); tabs.setPrefHeight(40);
        tabs.setStyle("-fx-open-tab-animation: none; -fx-close-tab-animation: none;");
        tabs.getSelectionModel().selectedIndexProperty().addListener((_, _, index) -> {
            if (!changingTabs && index.intValue() >= 0) { sheetIndex = index.intValue(); rowWindow = columnWindow = 0; refresh(); }
        });
        var foot = new HBox(20, status, selectionInfo); foot.setPadding(new Insets(8, 14, 8, 14));
        root.setBottom(new VBox(tabs, foot));
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
                    smokeSaveAndGit();
                } catch (Exception e) { failSmoke(e); }
            });
        }
    }
    private MenuBar menu() {
        var file = new Menu("파일");
        file.getItems().addAll(item("새 문서", this::newBook), item("열기 (.gsheet / .xlsx / .csv)", this::open),
                item("저장", () -> save(false)), item("다른 이름으로 저장", () -> save(true)), new SeparatorMenuItem(),
                item("Excel 내보내기 (.xlsx)", () -> export(false)), item("현재 시트 CSV 내보내기", () -> export(true)));
        var edit = new Menu("편집"); edit.getItems().addAll(item("실행 취소  Ctrl+Z", this::undo), item("다시 실행  Ctrl+Y", this::redo),
                item("복사  Ctrl+C", this::copy), item("잘라내기  Ctrl+X", () -> copy(true)), item("붙여넣기  Ctrl+V", this::paste),
                item("값만 붙여넣기", () -> paste(CellClipboard.Mode.VALUES)), item("서식만 붙여넣기", () -> paste(CellClipboard.Mode.FORMATS)),
                item("내용 지우기", this::clear), item("찾기", this::find), item("찾기 및 바꾸기", this::replaceAll),
                new SeparatorMenuItem(), item("아래로 채우기  Ctrl+D", () -> fill(SheetEdits.Direction.DOWN)), item("오른쪽 채우기  Ctrl+R", () -> fill(SheetEdits.Direction.RIGHT)));
        var format = new Menu("서식"); format.getItems().addAll(
                item("왼쪽 정렬", () -> applyStyle("align", "LEFT")), item("가운데 정렬", () -> applyStyle("align", "CENTER")), item("오른쪽 정렬", () -> applyStyle("align", "RIGHT")),
                item("모든 테두리", () -> applyStyle("border", "THIN")), item("테두리 없음", () -> applyStyle("border", "NONE")),
                item("글꼴 크기", () -> prompt("글꼴 크기", "크기 (6–72)", "13").ifPresent(value -> {
                    try { int size = Integer.parseInt(value); if (size < 6 || size > 72) throw new IllegalArgumentException("6–72 범위의 크기를 입력하세요."); applyStyle("size", value); }
                    catch (Exception e) { error(e); }
                })), item("열 너비", this::columnWidth));
        var sheet = new Menu("시트"); sheet.getItems().addAll(item("시트 추가", this::addSheet), item("시트 이름 변경", this::renameSheet),
                item("시트 삭제", this::deleteSheet), new SeparatorMenuItem(),
                item("선택 행 위에 삽입", () -> structuralEdit(StructuralEdits.Axis.ROW, true)),
                item("선택 행 전체 삭제", () -> structuralEdit(StructuralEdits.Axis.ROW, false)),
                item("선택 열 왼쪽에 삽입", () -> structuralEdit(StructuralEdits.Axis.COLUMN, true)),
                item("선택 열 전체 삭제", () -> structuralEdit(StructuralEdits.Axis.COLUMN, false)),
                new SeparatorMenuItem(), item("이전 1,000행", () -> page(-1)), item("다음 1,000행", () -> page(1)));
        var git = new Menu("Git"); git.getItems().addAll(item("서버 연결 / 동기화", this::gitConnect), item("문서 변경 내용 / 기록", this::gitInspect), item("문서 저장 후 커밋", this::gitCommit));
        var help = new Menu("도움말"); help.getItems().add(item("사용법 및 호환성", () -> textDialog("Git Sheet 0.2", "JDK 25 + JavaFX\n\n"
                + "더블 클릭 / F2: 셀 편집 · Enter: 적용 · Escape: 취소\n수식: =SUM(A1:A10), =IF(B1>0,\"예\",\"아니오\")\n"
                + "주소 상자: A1 또는 Z10000 입력 → 이동\n정렬·필터는 현재 1,000행 화면에만 적용되며 원본 행은 이동하지 않습니다.\n"
                + "Ctrl+C/V: 탭으로 구분된 직사각형 데이터 복사·붙여넣기\nCtrl+D/R: 아래로/오른쪽 채우기 (상대 참조 이동)\nCtrl+S: Git 친화적인 .gsheet 저장 · 저장 위치 알림 4초\n"
                + "Git → 서버 연결 / 동기화: GitHub·GitLab·Gitea·자체 서버 선택, push/fetch\n\n"
                + "지원: 값, POI 지원 수식, 기본 서식, 다중 시트, CSV/XLSX, 실행 취소, Git diff/commit, 간단한 차트\n"
                + "제한: Excel 완전 호환 제품이 아닙니다. VBA, Power Query, 피벗, 협업, 고급 차트 및 동적 배열 미지원.\n"
                + ".gsheet 변환은 그림, 이름 정의, 유효성 검사, 조건부 서식 등 고급 Excel 요소를 보존하지 않습니다.\n"
                + "병합·틀 고정 정보는 저장되지만 화면에는 반영되지 않습니다. 원본 .xlsx는 따로 보관하세요.\n"
                + "외부 연결은 갱신하지 않으며 지원하지 않는 수식은 #UNSUPPORTED!로 표시됩니다.")));
        return new MenuBar(file, edit, format, sheet, git, help);
    }
    private ToolBar toolbar() {
        var formats = new ComboBox<String>(FXCollections.observableArrayList("일반", "숫자", "통화", "백분율", "날짜"));
        formats.setValue("일반"); formats.setOnAction(_ -> applyStyle("format", switch (formats.getValue()) {
            case "숫자" -> "#,##0.00"; case "통화" -> "₩#,##0"; case "백분율" -> "0.00%"; case "날짜" -> "yyyy-mm-dd"; default -> "General";
        }));
        filter.setPromptText("현재 화면 행 필터"); filter.setPrefWidth(165); filter.setOnAction(_ -> refreshRows());
        return new ToolBar(button("새 문서", this::newBook), button("열기", this::open), button("저장", () -> save(false)),
                new Separator(), button("↶", this::undo), button("↷", this::redo), new Separator(),
                button("굵게", () -> toggleFont("bold")), button("기울임", () -> toggleFont("italic")),
                button("노란 배경", () -> applyStyle("fill", Short.toString(IndexedColors.LIGHT_YELLOW.getIndex()))), formats,
                new Separator(), button("A→Z", () -> sort(false)), button("Z→A", () -> sort(true)), filter,
                button("차트", this::chart), button("Git", this::gitConnect));
    }
    private HBox formulaBar() {
        address.setPrefWidth(100); address.setMaxWidth(100); address.setOnAction(_ -> goTo());
        formula.setPromptText("값 또는 =수식을 입력하세요"); HBox.setHgrow(formula, Priority.ALWAYS);
        formula.setOnAction(_ -> { var p = selected(); if (p != null) mutate(() -> book.set(sheetIndex, p.row(), p.column(), formula.getText())); });
        var bar = new HBox(10, address, new Label("fx"), formula); bar.setPadding(new Insets(9, 12, 9, 12)); return bar;
    }
    private static Button button(String text, Runnable action) { var b = new Button(text); b.setOnAction(_ -> action.run()); return b; }
    private static MenuItem item(String text, Runnable action) { var m = new MenuItem(text); m.setOnAction(_ -> action.run()); return m; }
    private void refreshTabs() {
        changingTabs = true; tabs.getTabs().clear();
        for (var sheet : book.workbook()) tabs.getTabs().add(new Tab(sheet.getSheetName()));
        sheetIndex = Math.clamp(sheetIndex, 0, book.workbook().getNumberOfSheets() - 1);
        tabs.getSelectionModel().select(sheetIndex); changingTabs = false;
    }
    private void refresh() {
        grid.getColumns().clear();
        var numbers = new TableColumn<Integer, String>("#"); numbers.setPrefWidth(64); numbers.setSortable(false); numbers.setEditable(false);
        numbers.setCellValueFactory(data -> new ReadOnlyStringWrapper(Integer.toString(data.getValue() + 1))); grid.getColumns().add(numbers);
        for (int c = columnWindow; c < Math.min(Book.MAX_COLUMNS, columnWindow + COLUMN_WINDOW); c++) {
            final int columnIndex = c;
            var column = new TableColumn<Integer, String>(CellReference.convertNumToColString(c));
            column.setUserData(c); column.setSortable(false);
            column.setPrefWidth(Math.max(95, book.workbook().getSheetAt(sheetIndex).getColumnWidthInPixels(c)));
            column.setCellValueFactory(data -> new ReadOnlyStringWrapper(book.display(sheetIndex, data.getValue(), columnIndex)));
            column.setCellFactory(_ -> new GridCell(columnIndex));
            grid.getColumns().add(column);
        }
        refreshRows(); title();
    }
    private void refreshRows() {
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
        return new Position(grid.getItems().get(selected.getRow()), (Integer) selected.getTableColumn().getUserData());
    }
    private List<Position> selectedPositions() {
        var positions = new ArrayList<Position>();
        for (var p : grid.getSelectionModel().getSelectedCells()) {
            if (p.getColumn() > 0 && p.getRow() >= 0 && p.getRow() < grid.getItems().size())
                positions.add(new Position(grid.getItems().get(p.getRow()), (Integer) p.getTableColumn().getUserData()));
        }
        return positions;
    }
    private void selectionChanged() {
        var p = selected(); if (p == null) return;
        address.setText(Book.address(p.row(), p.column())); formula.setText(book.input(sheetIndex, p.row(), p.column()));
        var values = selectedPositions(); double sum = 0; int count = 0;
        for (var cell : values) {
            try { sum += Double.parseDouble(book.raw(sheetIndex, cell.row(), cell.column())); count++; }
            catch (NumberFormatException ignored) { }
        }
        selectionInfo.setText("선택 %d셀 · 숫자 %d · 합계 %s".formatted(values.size(), count, Double.toString(sum)));
    }
    private boolean mutate(Runnable action) {
        try { book.transaction(_ -> action.run()); pendingCut = null; dirty = true; grid.refresh(); title(); selectionChanged(); return true; }
        catch (Exception e) { error(e); grid.refresh(); return false; }
    }
    private void undo() { if (book.undo()) { pendingCut = null; dirty = true; refreshTabs(); refresh(); } }
    private void redo() { if (book.redo()) { pendingCut = null; dirty = true; refreshTabs(); refresh(); } }
    private void applyStyle(String property, String value) {
        var positions = selectedPositions(); if (positions.isEmpty()) return;
        mutate(() -> positions.forEach(p -> book.style(sheetIndex, p.row(), p.column(), property, value)));
    }
    private void toggleFont(String property) {
        var p = selected(); if (p == null) return;
        var cell = book.cell(sheetIndex, p.row(), p.column(), false);
        var font = book.workbook().getFontAt(cell == null ? 0 : cell.getCellStyle().getFontIndex());
        applyStyle(property, Boolean.toString(!(property.equals("bold") ? font.getBold() : font.getItalic())));
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
            var focus = grid.getFocusModel().getFocusedCell(); grid.edit(focus.getRow(), focus.getTableColumn()); event.consume();
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
        if (!filter.getText().isBlank() || !clipboardViewIsContiguous(new Bounds(0, grid.getItems().size() - 1, b.firstCol(), b.lastCol()))) {
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
            mutate(() -> { for (int r = 0; r < rows.size(); r++) for (int c = 0; c < rows.get(r).size(); c++)
                book.set(sheetIndex, grid.getItems().get(start + r), firstCol + c,
                        mode == CellClipboard.Mode.VALUES ? "'" + rows.get(r).get(c) : rows.get(r).get(c)); });
        } catch (Exception e) { error(e); }
    }
    private void pasteCut(CutRange cut, int row, int column) {
        if (mutate(() -> CellMove.move(book, cut.sheet(), cut.row(), cut.column(), cut.rows(), cut.columns(), row, column))) {
            cellClipboard = null; copyId = null; status.setText("셀 이동 완료 · 참조 갱신 · Ctrl+Z로 실행 취소");
        }
    }
    private void clear() { var positions = selectedPositions(); if (!positions.isEmpty()) mutate(() -> positions.forEach(p -> book.set(sheetIndex, p.row(), p.column(), ""))); }
    private void goTo() {
        try {
            var ref = new CellReference(address.getText().trim().toUpperCase(Locale.ROOT));
            int r = ref.getRow(), c = ref.getCol(); book.cell(sheetIndex, r, c, false);
            rowWindow = (r / ROW_WINDOW) * ROW_WINDOW; columnWindow = (c / COLUMN_WINDOW) * COLUMN_WINDOW; filter.clear(); refresh();
            int index = r - rowWindow; grid.getSelectionModel().clearAndSelect(index, grid.getColumns().get(c - columnWindow + 1)); grid.scrollTo(index);
        } catch (Exception e) { error(new IllegalArgumentException("A1, Z10000처럼 올바른 셀 주소를 입력하세요.")); }
    }
    private void page(int direction) { rowWindow = Math.clamp(rowWindow + direction * ROW_WINDOW, 0, ((Book.MAX_ROWS - 1) / ROW_WINDOW) * ROW_WINDOW); refreshRows(); }
    private void sort(boolean descending) {
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
        if (!filter.getText().isBlank() || !grid.getSortOrder().isEmpty()) {
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
            sheetIndex = book.workbook().getNumberOfSheets() - 1; refreshTabs(); refresh();
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
            int deleting = sheetIndex; sheetIndex = 0;
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
        grid.edit(1, grid.getColumns().get(2)); grid.applyCss(); grid.layout();
        var editor = grid.lookupAll(".text-field").stream().filter(n -> n instanceof TextField && n.isVisible()).map(n -> (TextField)n).findFirst().orElseThrow();
        editor.setText("5"); editor.fireEvent(new javafx.event.ActionEvent());
        if (!book.raw(0, 1, 1).equals("5") || !book.display(0, 1, 3).equals("₩600,000")) throw new IllegalStateException("Edit/recalculation failed");
        undo(); if (!book.raw(0, 1, 1).equals("3")) throw new IllegalStateException("Undo failed");
        redo(); if (!book.raw(0, 1, 1).equals("5")) throw new IllegalStateException("Redo failed");
        undo();
        // Avoid changing the user's system clipboard during smoke checks.
        filter.setText("마우스"); refreshRows();
        if (grid.getItems().size() != 1 || grid.getItems().getFirst() != 2) throw new IllegalStateException("Filter failed");
        filter.clear(); address.setText("AZ2001"); goTo();
        if (rowWindow != 2000 || !Objects.equals(selected(), new Position(2000, 51))) throw new IllegalStateException("Navigation failed");
        address.setText("A1"); goTo(); grid.applyCss(); grid.layout();
        String original = book.raw(0, 0, 0);
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
    private void screenshot(Path path) throws IOException {
        var image = stage.getScene().getRoot().snapshot(null, null);
        var rendered = new java.awt.image.BufferedImage((int) image.getWidth(), (int) image.getHeight(), java.awt.image.BufferedImage.TYPE_INT_ARGB);
        var pixels = image.getPixelReader();
        for (int y = 0; y < rendered.getHeight(); y++) for (int x = 0; x < rendered.getWidth(); x++) rendered.setRGB(x, y, pixels.getArgb(x, y));
        if (path.toAbsolutePath().getParent() != null) Files.createDirectories(path.toAbsolutePath().getParent());
        javax.imageio.ImageIO.write(rendered, "png", path.toFile());
    }
    @SuppressWarnings("unchecked")
    private void smokeSaveAndGit() throws Exception {
        var folder = Files.createTempDirectory(Path.of("target"), "ui-smoke-").toAbsolutePath();
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
        private TextField editor;
        GridCell(int column) { this.column = column; }
        @Override protected void updateItem(String value, boolean empty) {
            super.updateItem(value, empty);
            if (empty) { setText(null); setGraphic(null); setStyle(""); return; }
            if (!isEditing()) { setText(value); setGraphic(null); }
            int row = getTableRow().getItem() == null ? -1 : getTableRow().getItem();
            var cell = row < 0 ? null : book.cell(sheetIndex, row, column, false);
            if (cell == null) { setStyle(""); return; }
            var style = cell.getCellStyle(); var font = book.workbook().getFontAt(style.getFontIndex());
            var css = new StringBuilder("-fx-font-weight:").append(font.getBold() ? "bold" : "normal").append(";-fx-font-style:").append(font.getItalic() ? "italic" : "normal").append(';');
            if (style instanceof org.apache.poi.xssf.usermodel.XSSFCellStyle xs && style.getFillPattern() == FillPatternType.SOLID_FOREGROUND && !isSelected()) {
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
            setStyle(css.toString());
        }
        @Override public void startEdit() {
            if (isEmpty()) return;
            super.startEdit(); int row = getTableRow().getItem();
            editor = new TextField(book.input(sheetIndex, row, column));
            pendingEdit = () -> finish(row);
            editor.setOnAction(_ -> finish(row)); editor.setOnKeyPressed(event -> { if (event.getCode() == KeyCode.ESCAPE) { cancelEdit(); event.consume(); } });
            editor.focusedProperty().addListener((_, _, focus) -> { if (!focus && isEditing()) finish(row); });
            setText(null); setGraphic(editor); editor.requestFocus(); editor.selectAll();
        }
        private boolean finish(int row) {
            if (!isEditing()) { pendingEdit = null; return true; }
            String text = editor.getText(); pendingEdit = null; super.commitEdit(text); setGraphic(null);
            return mutate(() -> book.set(sheetIndex, row, column, text));
        }
        @Override public void cancelEdit() { pendingEdit = null; super.cancelEdit(); setGraphic(null); setText(getItem()); }
    }
    @Override public void stop() throws Exception { background.shutdownNow(); book.close(); }
}
