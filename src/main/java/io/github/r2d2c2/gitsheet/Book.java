package io.github.r2d2c2.gitsheet;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
import java.util.function.Consumer;

/** UI-thread-confined workbook. POI handles Excel reference semantics and evaluation. */
public final class Book implements AutoCloseable {
    public static final int MAX_ROWS = 1_048_576, MAX_COLUMNS = 16_384;
    private XSSFWorkbook workbook;
    private FormulaEvaluator evaluator;
    private final DataFormatter formatter = new DataFormatter(Locale.KOREA);
    private final Deque<byte[]> undo = new ArrayDeque<>(), redo = new ArrayDeque<>();
    private static final int HISTORY_LIMIT = 30;
    private static final long HISTORY_BYTES = 32L * 1024 * 1024;
    private final Map<String, CellStyle> styleCache = new HashMap<>();

    public Book() { this(new XSSFWorkbook()); workbook.createSheet("Sheet1"); }
    public Book(XSSFWorkbook workbook) { this.workbook = workbook; resetEvaluator(); }
    public XSSFWorkbook workbook() { return workbook; }
    public void resetEvaluator() { evaluator = workbook.getCreationHelper().createFormulaEvaluator(); styleCache.clear(); }
    public Cell cell(int sheet, int row, int col, boolean create) {
        if (row < 0 || row >= MAX_ROWS || col < 0 || col >= MAX_COLUMNS) throw new IllegalArgumentException("셀 범위를 벗어났습니다.");
        var s = workbook.getSheetAt(sheet);
        var r = s.getRow(row);
        if (r == null && create) r = s.createRow(row);
        return r == null ? null : (create ? r.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK) : r.getCell(col));
    }
    public String raw(int sheet, int row, int col) {
        var cell = cell(sheet, row, col, false);
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case FORMULA -> "=" + cell.getCellFormula();
            case NUMERIC -> org.apache.poi.ss.util.NumberToTextConverter.toText(cell.getNumericCellValue());
            case BOOLEAN -> Boolean.toString(cell.getBooleanCellValue()).toUpperCase(Locale.ROOT);
            case STRING -> cell.getStringCellValue();
            case ERROR -> FormulaError.forInt(cell.getErrorCellValue()).getString();
            default -> "";
        };
    }
    public String input(int sheet, int row, int col) {
        var cell = cell(sheet, row, col, false);
        String value = raw(sheet, row, col);
        return cell != null && cell.getCellType() == CellType.STRING ? "'" + value : value;
    }
    public String display(int sheet, int row, int col) {
        var cell = cell(sheet, row, col, false);
        if (cell == null) return "";
        try { return formatter.formatCellValue(cell, evaluator); }
        catch (RuntimeException e) { return "#UNSUPPORTED!"; }
    }
    public void set(int sheet, int row, int col, String text) {
        var c = cell(sheet, row, col, true);
        c.setBlank();
        if (text.startsWith("=")) c.setCellFormula(text.substring(1));
        else if (text.startsWith("'")) c.setCellValue(text.substring(1));
        else if (text.equalsIgnoreCase("true") || text.equalsIgnoreCase("false")) c.setCellValue(Boolean.parseBoolean(text));
        else if (!text.isEmpty()) {
            try {
                double number = Double.parseDouble(text);
                if (Double.isFinite(number) && !text.matches("[+-]?0\\d+.*")) c.setCellValue(number);
                else c.setCellValue(text);
            } catch (NumberFormatException e) { c.setCellValue(text); }
        }
        evaluator.clearAllCachedResultValues();
    }
    public void transaction(Consumer<Book> change) {
        var before = snapshot();
        try { change.accept(this); }
        catch (RuntimeException failure) { restore(before); throw failure; }
        undo.push(before); trim(undo); redo.clear(); evaluator.clearAllCachedResultValues();
    }
    private static void trim(Deque<byte[]> history) {
        long bytes = history.stream().mapToLong(b -> b.length).sum();
        while (history.size() > HISTORY_LIMIT || (bytes > HISTORY_BYTES && history.size() > 1)) bytes -= history.removeLast().length;
    }
    public boolean undo() {
        if (undo.isEmpty()) return false;
        redo.push(snapshot()); trim(redo); restore(undo.pop()); return true;
    }
    public boolean redo() {
        if (redo.isEmpty()) return false;
        undo.push(snapshot()); trim(undo); restore(redo.pop()); return true;
    }
    public byte[] snapshot() {
        try (var bytes = new ByteArrayOutputStream()) { workbook.write(bytes); return bytes.toByteArray(); }
        catch (IOException e) { throw new UncheckedIOException(e); }
    }
    private void restore(byte[] bytes) {
        try { var next = new XSSFWorkbook(new ByteArrayInputStream(bytes)); workbook.close(); workbook = next; resetEvaluator(); }
        catch (IOException e) { throw new UncheckedIOException(e); }
    }
    public void style(int sheet, int row, int col, String property, String value) {
        var c = cell(sheet, row, col, true);
        var old = c.getCellStyle();
        String key = old.getIndex() + ":" + property + ":" + value;
        var style = styleCache.computeIfAbsent(key, ignored -> {
            var result = workbook.createCellStyle(); result.cloneStyleFrom(old);
            switch (property) {
                case "format" -> result.setDataFormat(workbook.createDataFormat().getFormat(value));
                case "align" -> result.setAlignment(HorizontalAlignment.valueOf(value));
                case "wrap" -> result.setWrapText(Boolean.parseBoolean(value));
                case "fill" -> { result.setFillForegroundColor(Short.parseShort(value)); result.setFillPattern(FillPatternType.SOLID_FOREGROUND); }
                case "border" -> {
                    var border = BorderStyle.valueOf(value); result.setBorderTop(border); result.setBorderBottom(border); result.setBorderLeft(border); result.setBorderRight(border);
                }
                case "bold", "italic", "size", "underline" -> {
                    var previous = workbook.getFontAt(old.getFontIndex());
                    var font = workbook.createFont(); font.getCTFont().set(previous.getCTFont());
                    switch (property) {
                        case "bold" -> font.setBold(Boolean.parseBoolean(value));
                        case "italic" -> font.setItalic(Boolean.parseBoolean(value));
                        case "size" -> font.setFontHeightInPoints(Short.parseShort(value));
                        case "underline" -> font.setUnderline(Boolean.parseBoolean(value) ? Font.U_SINGLE : Font.U_NONE);
                        default -> throw new IllegalArgumentException(property);
                    }
                    result.setFont(font);
                }
                default -> throw new IllegalArgumentException(property);
            }
            return result;
        });
        c.setCellStyle(style);
    }
    public static String address(int row, int col) { return CellReference.convertNumToColString(col) + (row + 1); }
    @Override public void close() throws IOException { workbook.close(); }
}
