package io.github.r2d2c2.gitsheet;

import java.io.*;
import java.util.HashMap;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Immutable copy-time snapshot, independent of later edits and undo workbook replacement. */
public final class CellClipboard {
    public enum Mode { ALL, VALUES, FORMATS }
    private final byte[] snapshot;
    private final int sheet, top, left, rows, columns;

    public CellClipboard(Book book, int sheet, int top, int left, int rows, int columns) {
        if (rows < 1 || columns < 1) throw new IllegalArgumentException("복사 범위가 비어 있습니다.");
        book.cell(sheet, top, left, false);
        book.cell(sheet, Math.addExact(top, rows - 1), Math.addExact(left, columns - 1), false);
        this.sheet = sheet; this.top = top; this.left = left; this.rows = rows; this.columns = columns;
        snapshot = book.snapshot();
    }

    /** Caller uses Book.transaction so a failed paste rolls back and a successful paste is one undo. */
    public void paste(Book target, int targetSheet, int row, int column, Mode mode) {
        target.cell(targetSheet, row, column, false);
        target.cell(targetSheet, Math.addExact(row, rows - 1), Math.addExact(column, columns - 1), false);
        try (var source = new Book(new XSSFWorkbook(new ByteArrayInputStream(snapshot)))) {
            var styles = new HashMap<Integer, CellStyle>();
            var evaluator = source.workbook().getCreationHelper().createFormulaEvaluator();
            for (int r = 0; r < rows; r++) for (int c = 0; c < columns; c++) {
                var original = source.cell(sheet, top + r, left + c, false);
                var dest = target.cell(targetSheet, row + r, column + c, true);
                if (mode != Mode.VALUES) {
                    int index = original == null ? 0 : Short.toUnsignedInt(original.getCellStyle().getIndex());
                    dest.setCellStyle(styles.computeIfAbsent(index, key -> {
                        var style = target.workbook().createCellStyle();
                        style.cloneStyleFrom(source.workbook().getCellStyleAt(key)); return style;
                    }));
                }
                if (mode == Mode.FORMATS) continue;
                dest.setBlank();
                if (original == null) continue;
                if (original.getCellType() == CellType.FORMULA && mode == Mode.ALL) {
                    dest.setCellFormula(SheetEdits.translate(source, sheet, original.getCellFormula(), top + r, left + c, row + r, column + c));
                    continue;
                }
                var value = evaluator.evaluate(original);
                if (value == null) continue;
                switch (value.getCellType()) {
                    case STRING -> dest.setCellValue(value.getStringValue());
                    case NUMERIC -> dest.setCellValue(value.getNumberValue());
                    case BOOLEAN -> dest.setCellValue(value.getBooleanValue());
                    case ERROR -> dest.setCellErrorValue(value.getErrorValue());
                    default -> dest.setBlank();
                }
            }
            target.resetEvaluator();
        } catch (IOException e) { throw new UncheckedIOException(e); }
    }
}
