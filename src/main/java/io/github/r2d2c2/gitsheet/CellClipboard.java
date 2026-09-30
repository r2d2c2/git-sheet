package io.github.r2d2c2.gitsheet;

import java.io.*;
import java.util.*;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.*;
import org.apache.poi.ss.formula.ptg.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

/** Copy-time range snapshot. Holds no source workbook, cells, or unselected data. */
public final class CellClipboard {
    public enum Mode { ALL, VALUES, FORMATS }
    private record Entry(CellType type, Object value, int style, Formula formula, CellValue calculated, String calculationFailure) {}
    private final byte[] styleSnapshot;
    private final List<Entry> entries;
    private final int sheet, top, left, rows, columns;
    private final String sheetName;

    public CellClipboard(Book book, int sheet, int top, int left, int rows, int columns) {
        if (rows < 1 || columns < 1) throw new IllegalArgumentException("복사 범위가 비어 있습니다.");
        book.cell(sheet, top, left, false);
        book.cell(sheet, Math.addExact(top, rows - 1), Math.addExact(left, columns - 1), false);
        this.sheet = sheet; this.top = top; this.left = left; this.rows = rows; this.columns = columns;
        sheetName = book.workbook().getSheetName(sheet);
        var evaluator = book.workbook().getCreationHelper().createFormulaEvaluator();
        var parser = XSSFEvaluationWorkbook.create(book.workbook());
        var captured = new ArrayList<Entry>();
        try (var styles = new XSSFWorkbook(); var bytes = new ByteArrayOutputStream()) {
            styles.createSheet("Styles");
            var styleIndexes = new HashMap<Integer, Integer>();
            for (int r = 0; r < rows; r++) for (int c = 0; c < columns; c++) {
                var cell = book.cell(sheet, top + r, left + c, false);
                int originalStyle = cell == null ? 0 : Short.toUnsignedInt(cell.getCellStyle().getIndex());
                int style = styleIndexes.computeIfAbsent(originalStyle, key -> {
                    var copy = styles.createCellStyle(); copy.cloneStyleFrom(book.workbook().getCellStyleAt(key));
                    return Short.toUnsignedInt(copy.getIndex());
                });
                var type = cell == null ? CellType.BLANK : cell.getCellType();
                Object value = switch (type) {
                    case STRING -> cell.getStringCellValue(); case NUMERIC -> cell.getNumericCellValue();
                    case BOOLEAN -> cell.getBooleanCellValue(); case ERROR -> cell.getErrorCellValue();
                    default -> null;
                };
                Formula formula = null; CellValue calculated = null; String failure = null;
                if (type == CellType.FORMULA) {
                    formula = Formula.capture(cell.getCellFormula(), parser, sheet, top + r);
                    try { calculated = evaluator.evaluate(cell); }
                    catch (RuntimeException e) { failure = "수식 계산을 지원하지 않아 값만 붙여넣을 수 없습니다: " + cell.getAddress(); }
                }
                captured.add(new Entry(type, value, style, formula, calculated, failure));
            }
            styles.write(bytes); styleSnapshot = bytes.toByteArray(); entries = List.copyOf(captured);
        } catch (IOException e) { throw new UncheckedIOException(e); }
    }

    /** Caller uses Book.transaction for one-step undo and rollback on unsupported formulas. */
    public void paste(Book target, int targetSheet, int row, int column, Mode mode) {
        target.cell(targetSheet, row, column, false);
        target.cell(targetSheet, Math.addExact(row, rows - 1), Math.addExact(column, columns - 1), false);
        if (mode == Mode.VALUES) for (var entry : entries)
            if (entry.calculationFailure() != null) throw new IllegalArgumentException(entry.calculationFailure());
        try (var styleBook = mode == Mode.VALUES ? null : new XSSFWorkbook(new ByteArrayInputStream(styleSnapshot))) {
            var styles = new HashMap<Integer, CellStyle>();
            int index = 0;
            for (int r = 0; r < rows; r++) for (int c = 0; c < columns; c++) {
                var entry = entries.get(index++);
                var dest = target.cell(targetSheet, row + r, column + c, true);
                if (mode != Mode.VALUES) dest.setCellStyle(styles.computeIfAbsent(entry.style(), key -> {
                    var style = target.workbook().createCellStyle(); style.cloneStyleFrom(styleBook.getCellStyleAt(key)); return style;
                }));
                if (mode == Mode.FORMATS) continue;
                dest.setBlank();
                if (entry.type() == CellType.FORMULA) {
                    if (mode == Mode.ALL) dest.setCellFormula(entry.formula().translate(sheet, sheetName, top + r, left + c, row + r, column + c));
                    else writeCalculated(dest, entry.calculated());
                } else switch (entry.type()) {
                    case STRING -> dest.setCellValue((String) entry.value()); case NUMERIC -> dest.setCellValue((Double) entry.value());
                    case BOOLEAN -> dest.setCellValue((Boolean) entry.value()); case ERROR -> dest.setCellErrorValue((Byte) entry.value());
                    default -> { }
                }
            }
            target.resetEvaluator();
        } catch (IOException e) { throw new UncheckedIOException(e); }
    }
    private static void writeCalculated(Cell cell, CellValue value) {
        if (value == null) return;
        switch (value.getCellType()) {
            case STRING -> cell.setCellValue(value.getStringValue()); case NUMERIC -> cell.setCellValue(value.getNumberValue());
            case BOOLEAN -> cell.setCellValue(value.getBooleanValue()); case ERROR -> cell.setCellErrorValue(value.getErrorValue());
            default -> { }
        }
    }

    /** Capture only names/sheet identifiers actually requested when rendering these tokens. */
    private record Formula(Ptg[] tokens, FormulaRenderingWorkbook rendering, String failure) {
        static Formula capture(String text, XSSFEvaluationWorkbook book, int sheet, int row) {
            try { return parse(text, book, sheet, row); }
            catch (RuntimeException e) { return new Formula(null, null, "수식 참조 변환을 지원하지 않습니다: " + text); }
        }
        private static Formula parse(String text, XSSFEvaluationWorkbook book, int sheet, int row) {
            var tokens = FormulaParser.parse(text, book, FormulaType.CELL, sheet, row);
            var sheets = new HashMap<Integer, EvaluationWorkbook.ExternalSheet>();
            var first = new HashMap<Integer, String>(); var last = new HashMap<Integer, String>();
            var names = new HashMap<Integer, String>(); var externalNames = new HashMap<String, String>();
            var recorder = new FormulaRenderingWorkbook() {
                public EvaluationWorkbook.ExternalSheet getExternalSheet(int i) { var v = book.getExternalSheet(i); sheets.put(i, v); return v; }
                public String getSheetFirstNameByExternSheet(int i) { var v = book.getSheetFirstNameByExternSheet(i); first.put(i, v); return v; }
                public String getSheetLastNameByExternSheet(int i) { var v = book.getSheetLastNameByExternSheet(i); last.put(i, v); return v; }
                public String getNameText(NamePtg t) { var v = book.getNameText(t); names.put(t.getIndex(), v); return v; }
                public String resolveNameXText(NameXPtg t) { var v = book.resolveNameXText(t); externalNames.put(key(t), v); return v; }
            };
            FormulaRenderer.toFormulaString(recorder, tokens);
            return new Formula(tokens, new Rendering(sheets, first, last, names, externalNames), null);
        }
        String translate(int sheet, String name, int fromRow, int fromColumn, int toRow, int toColumn) {
            if (failure != null) throw new IllegalArgumentException(failure);
            var moved = Arrays.stream(tokens).map(Ptg::copy).toArray(Ptg[]::new);
            if (fromRow != toRow) FormulaShifter.createForRowCopy(sheet, name, fromRow, fromRow, toRow - fromRow, SpreadsheetVersion.EXCEL2007).adjustFormula(moved, sheet);
            if (fromColumn != toColumn) FormulaShifter.createForColumnCopy(sheet, name, fromColumn, fromColumn, toColumn - fromColumn, SpreadsheetVersion.EXCEL2007).adjustFormula(moved, sheet);
            return FormulaRenderer.toFormulaString(rendering, moved);
        }
    }
    private static String key(NameXPtg t) { return t.getSheetRefIndex() + ":" + t.getNameIndex(); }
    private record Rendering(Map<Integer, EvaluationWorkbook.ExternalSheet> sheets, Map<Integer, String> first,
                             Map<Integer, String> last, Map<Integer, String> names, Map<String, String> externalNames) implements FormulaRenderingWorkbook {
        public EvaluationWorkbook.ExternalSheet getExternalSheet(int i) { return sheets.get(i); }
        public String getSheetFirstNameByExternSheet(int i) { return first.get(i); }
        public String getSheetLastNameByExternSheet(int i) { return last.get(i); }
        public String getNameText(NamePtg t) { return names.get(t.getIndex()); }
        public String resolveNameXText(NameXPtg t) { return externalNames.get(key(t)); }
    }
}
