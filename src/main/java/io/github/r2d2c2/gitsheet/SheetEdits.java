package io.github.r2d2c2.gitsheet;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import java.util.regex.*;

/** Range editing primitives; callers wrap operations in Book.transaction for atomic undo. */
public final class SheetEdits {
    private SheetEdits() {}
    public enum Direction { DOWN, RIGHT }
    public static void fill(Book book, int sheet, int top, int bottom, int left, int right, Direction direction) {
        if (top > bottom || left > right) throw new IllegalArgumentException("잘못된 선택 범위입니다.");
        book.cell(sheet, top, left, false); book.cell(sheet, bottom, right, false);
        checkWritableRange(book, sheet, top, bottom, left, right);
        for (int r = top; r <= bottom; r++) for (int c = left; c <= right; c++) {
            int sourceRow = direction == Direction.DOWN ? top : r;
            int sourceColumn = direction == Direction.RIGHT ? left : c;
            if (sourceRow == r && sourceColumn == c) continue;
            var source = book.cell(sheet, sourceRow, sourceColumn, false);
            var target = book.cell(sheet, r, c, true); target.setBlank();
            if (source == null) { target.setCellStyle(book.workbook().getCellStyleAt(0)); continue; }
            target.setCellStyle(source.getCellStyle());
            switch (source.getCellType()) {
                case STRING -> target.setCellValue(source.getStringCellValue());
                case NUMERIC -> target.setCellValue(source.getNumericCellValue());
                case BOOLEAN -> target.setCellValue(source.getBooleanCellValue());
                case ERROR -> target.setCellErrorValue(source.getErrorCellValue());
                case FORMULA -> target.setCellFormula(translate(book, sheet, source.getCellFormula(), sourceRow, sourceColumn, r, c));
                default -> target.setBlank();
            }
        }
        book.resetEvaluator();
    }
    /** Preflight all destinations before a bulk operation changes cells or allocates styles. */
    public static void checkWritableRange(Book book, int sheet, int top, int bottom, int left, int right) {
        if (top > bottom || left > right) throw new IllegalArgumentException("잘못된 선택 범위입니다.");
        book.cell(sheet, top, left, false); book.cell(sheet, bottom, right, false);
        var targetSheet = book.workbook().getSheetAt(sheet);
        if (targetSheet.getProtect()) throw new IllegalArgumentException("보호된 시트에서는 변경할 수 없습니다.");
        var range = new org.apache.poi.ss.util.CellRangeAddress(top, bottom, left, right);
        for (var merged : targetSheet.getMergedRegions()) {
            if (merged.intersects(range)) throw new IllegalArgumentException("병합 셀을 포함한 범위는 병합을 해제한 뒤 수정해 주세요.");
        }
        // Validate the entire range before touching any destination, including array sources.
        for (int index = top; index <= bottom; index++) {
            var row = targetSheet.getRow(index);
            if (row == null) continue;
            for (var cell : row) {
                if (cell.getColumnIndex() >= left && cell.getColumnIndex() <= right && cell.isPartOfArrayFormulaGroup())
                    throw new IllegalArgumentException("배열 수식을 포함한 범위는 이 범위 편집을 지원하지 않습니다.");
            }
        }
    }
    /** Ragged external clipboard rows: validate every written cell before writing the first row. */
    public static void pasteText(Book book, int sheet, int top, int left, java.util.List<java.util.List<String>> rows, boolean literal) {
        for (int r = 0; r < rows.size(); r++) {
            var values = rows.get(r);
            if (!values.isEmpty()) checkWritableRange(book, sheet, Math.addExact(top, r), Math.addExact(top, r), left,
                    Math.addExact(left, values.size() - 1));
        }
        for (int r = 0; r < rows.size(); r++) for (int c = 0; c < rows.get(r).size(); c++)
            book.set(sheet, top + r, left + c, literal ? "'" + rows.get(r).get(c) : rows.get(r).get(c));
    }
    public static String translate(Book book, int sheet, String formula, int fromRow, int fromColumn, int toRow, int toColumn) {
        var evaluator = XSSFEvaluationWorkbook.create(book.workbook());
        var tokens = FormulaParser.parse(formula, evaluator, FormulaType.CELL, sheet, fromRow);
        String name = book.workbook().getSheetName(sheet);
        if (fromRow != toRow) FormulaShifter.createForRowCopy(sheet, name, fromRow, fromRow, toRow - fromRow, SpreadsheetVersion.EXCEL2007).adjustFormula(tokens, sheet);
        if (fromColumn != toColumn) FormulaShifter.createForColumnCopy(sheet, name, fromColumn, fromColumn, toColumn - fromColumn, SpreadsheetVersion.EXCEL2007).adjustFormula(tokens, sheet);
        return FormulaRenderer.toFormulaString(evaluator, tokens);
    }
    public static int replaceAll(Book book, int sheet, String find, String replacement, boolean matchCase, boolean formulas) {
        if (find.isEmpty()) throw new IllegalArgumentException("찾을 내용을 입력하세요.");
        var pattern = Pattern.compile(Pattern.quote(find), matchCase ? 0 : Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE);
        int changed = 0;
        for (var row : book.workbook().getSheetAt(sheet)) for (var cell : row) {
            boolean formula = cell.getCellType() == CellType.FORMULA;
            if (cell.getCellType() != CellType.STRING && !(formulas && formula)) continue;
            String text = formula ? cell.getCellFormula() : cell.getStringCellValue();
            String next = pattern.matcher(text).replaceAll(Matcher.quoteReplacement(replacement));
            if (!text.equals(next)) {
                if (formula) cell.setCellFormula(next); else cell.setCellValue(next);
                changed++;
            }
        }
        book.resetEvaluator(); return changed;
    }
}
