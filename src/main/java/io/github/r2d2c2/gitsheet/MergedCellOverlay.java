package io.github.r2d2c2.gitsheet;

import java.util.*;
import java.util.function.*;
import javafx.application.Platform;
import javafx.geometry.Insets;
import javafx.scene.control.*;
import javafx.scene.input.*;
import javafx.scene.layout.Pane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellAddress;

/** Virtualized merged-cell presentation. All access is confined to the JavaFX thread. */
public final class MergedCellOverlay implements AutoCloseable {
    private final TableView<Integer> grid;
    private final Pane mergeLayer;
    private final Supplier<Book> bookSupplier;
    private final IntSupplier sheetSupplier;
    private final Supplier<CellAddress> selection;
    private final BiConsumer<CellAddress, Boolean> activate;
    private final BiFunction<Cell, Boolean, String> style;
    private final Map<String, Label> mergeLabels = new HashMap<>();
    private boolean pending, closed;

    public MergedCellOverlay(TableView<Integer> grid, Pane layer, Supplier<Book> book, IntSupplier sheet,
            Supplier<CellAddress> selection, BiConsumer<CellAddress, Boolean> activate,
            BiFunction<Cell, Boolean, String> style) {
        this.grid = grid; mergeLayer = layer; bookSupplier = book; sheetSupplier = sheet;
        this.selection = selection; this.activate = activate; this.style = style;
        grid.needsLayoutProperty().addListener((_, _, needed) -> { if (needed) requestPaint(); });
        grid.addEventHandler(ScrollEvent.SCROLL, _ -> requestPaint());
        grid.addEventHandler(MouseEvent.MOUSE_DRAGGED, _ -> requestPaint());
        grid.editingCellProperty().addListener((_, _, _) -> requestPaint());
        layer.widthProperty().addListener((_, _, _) -> requestPaint());
        layer.heightProperty().addListener((_, _, _) -> requestPaint());
    }
    public void requestPaint() {
        if (closed || pending) return;
        pending = true;
        Platform.runLater(() -> { try { paint(); } finally { pending = false; } });
    }
    public void paint() {
        if (closed || grid.getScene() == null) return;
        var book = bookSupplier.get();
        int sheetIndex = sheetSupplier.getAsInt();
        var sheet = book.workbook().getSheetAt(sheetIndex);
        var ranges = sheet.getMergedRegions();
        if (ranges.isEmpty()) { mergeLayer.getChildren().clear(); mergeLabels.clear(); return; }
        grid.applyCss(); grid.layout();
        var flow = grid.lookup(".virtual-flow");
        if (flow == null) return;
        var viewport = mergeLayer.sceneToLocal(flow.localToScene(flow.getBoundsInLocal()));
        double right = viewport.getMaxX(), bottom = viewport.getMaxY();
        for (var node : grid.lookupAll(".scroll-bar")) if (node instanceof ScrollBar bar && bar.isVisible()) {
            var bound = mergeLayer.sceneToLocal(bar.localToScene(bar.getBoundsInLocal()));
            if (bar.getOrientation() == javafx.geometry.Orientation.VERTICAL) right = Math.min(right, bound.getMinX());
            else bottom = Math.min(bottom, bound.getMinY());
        }
        var clip = mergeLayer.getClip() instanceof javafx.scene.shape.Rectangle rectangle ? rectangle : new javafx.scene.shape.Rectangle();
        clip.setX(viewport.getMinX()); clip.setY(viewport.getMinY());
        clip.setWidth(Math.max(0, right - viewport.getMinX())); clip.setHeight(Math.max(0, bottom - viewport.getMinY()));
        if (mergeLayer.getClip() != clip) mergeLayer.setClip(clip);
        var visible = grid.lookupAll(".table-cell").stream().filter(n -> n instanceof TableCell<?, ?> cell && !cell.isEmpty()
                && cell.getTableRow() != null && cell.getTableRow().getItem() instanceof Integer && cell.getTableColumn().getUserData() instanceof Integer).map(n -> (TableCell<?, ?>)n).toList();
        var active = selection.get();
        var children = new ArrayList<javafx.scene.Node>(); var keys = new HashSet<String>();
        for (var range : ranges) {
            double x = Double.POSITIVE_INFINITY, y = Double.POSITIVE_INFINITY, x2 = Double.NEGATIVE_INFINITY, y2 = Double.NEGATIVE_INFINITY;
            boolean editing = false;
            for (var cell : visible) if (range.isInRange((Integer) cell.getTableRow().getItem(), (Integer) cell.getTableColumn().getUserData())) {
                editing |= cell.isEditing();
                var b = mergeLayer.sceneToLocal(cell.localToScene(cell.getBoundsInLocal()));
                x = Math.min(x, b.getMinX()); y = Math.min(y, b.getMinY()); x2 = Math.max(x2, b.getMaxX()); y2 = Math.max(y2, b.getMaxY());
            }
            if (editing || !Double.isFinite(x) || x2 <= viewport.getMinX() || y2 <= viewport.getMinY() || x >= right || y >= bottom) continue;
            String key = range.formatAsString(); keys.add(key);
            var label = mergeLabels.computeIfAbsent(key, _ -> {
                var result = new Label(); result.setManaged(false); result.setPadding(new Insets(4));
                result.setOnMousePressed(event -> {
                    if (event.getButton() == MouseButton.PRIMARY) {
                        activate.accept(new CellAddress(range.getFirstRow(), range.getFirstColumn()), event.getClickCount() >= 2); event.consume();
                    }
                });
                return result;
            });
            var cell = book.cell(sheetIndex, range.getFirstRow(), range.getFirstColumn(), false);
            label.setText(book.display(sheetIndex, range.getFirstRow(), range.getFirstColumn()));
            label.setWrapText(cell != null && cell.getCellStyle().getWrapText());
            boolean selected = active != null && range.isInRange(active.getRow(), active.getColumn());
            label.setStyle("-fx-background-color:white;-fx-text-fill:#172033;-fx-alignment:CENTER-LEFT;" + style.apply(cell, selected)
                    + (selected ? "-fx-background-color:#dbeafe;-fx-border-color:#2563eb;" : "-fx-border-color:#cbd5e1;") + "-fx-border-width:1;");
            label.resizeRelocate(x, y, x2-x, y2-y); children.add(label);
        }
        mergeLabels.keySet().retainAll(keys);
        if (!mergeLayer.getChildren().equals(children)) mergeLayer.getChildren().setAll(children);
    }
    /** Disable queued repaint work before the owning application closes its workbook. */
    @Override public void close() {
        closed = true;
        mergeLayer.getChildren().clear(); mergeLabels.clear();
    }
}
