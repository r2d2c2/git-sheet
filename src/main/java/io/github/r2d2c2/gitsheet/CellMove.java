package io.github.r2d2c2.gitsheet;

import java.util.*;
import org.apache.poi.ss.formula.*;
import org.apache.poi.ss.formula.ptg.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;

/** Same-sheet cell moves. Call inside Book.transaction for atomic rollback and undo. */
public final class CellMove {
    private CellMove() {}
    private record Value(CellType type, Object content, CellStyle style) {}

    public static void move(Book book, int sheet, int top, int left, int rows, int columns, int toRow, int toColumn) {
        if (rows < 1 || columns < 1) throw new IllegalArgumentException("이동 범위가 비어 있습니다.");
        book.cell(sheet, top, left, false); book.cell(sheet, toRow, toColumn, false);
        book.cell(sheet, Math.addExact(top, rows - 1), Math.addExact(left, columns - 1), false);
        book.cell(sheet, Math.addExact(toRow, rows - 1), Math.addExact(toColumn, columns - 1), false);
        if (top == toRow && left == toColumn) return;
        var source = new CellRangeAddress(top, top + rows - 1, left, left + columns - 1);
        var target = new CellRangeAddress(toRow, toRow + rows - 1, toColumn, toColumn + columns - 1);
        var workbook = book.workbook(); var active = workbook.getSheetAt(sheet);
        if (active.getProtect()) unsupported("보호된 시트");
        for (var merge : active.getMergedRegions()) if (merge.intersects(source) || merge.intersects(target)) unsupported("병합 셀");
        // These objects can contain additional formula dependencies or anchored content.
        // Until they can be updated together, refuse the move before changing any cells.
        for (var s : workbook) {
            if (!((org.apache.poi.xssf.usermodel.XSSFSheet) s).getTables().isEmpty() || !s.getDataValidations().isEmpty()
                    || s.getSheetConditionalFormatting().getNumConditionalFormattings() > 0 || s.getDrawingPatriarch() != null)
                unsupported("표·유효성 검사·조건부 서식·그림이 있는 문서");
            for (var row : s) for (var cell : row) {
                if (cell.isPartOfArrayFormulaGroup()) unsupported("배열 수식");
                if (s == active && (source.isInRange(cell) || target.isInRange(cell))
                        && (cell.getCellComment() != null || cell.getHyperlink() != null)) unsupported("메모·하이퍼링크 셀");
            }
        }
        var evaluation = XSSFEvaluationWorkbook.create(workbook);
        var formulas = new LinkedHashMap<Cell, String>();
        var names = new LinkedHashMap<Name, String>();
        int dr = toRow - top, dc = toColumn - left;
        for (var s : workbook) for (var row : s) for (var cell : row) if (cell.getCellType() == CellType.FORMULA) {
            // A formula overwritten by the destination does not survive this move.
            if (s == active && target.isInRange(cell) && !source.isInRange(cell)) continue;
            String before = cell.getCellFormula();
            String after = retarget(before, evaluation, FormulaType.CELL, workbook.getSheetIndex(s), cell.getRowIndex(), sheet, active.getSheetName(), source, dr, dc);
            if (!before.equals(after)) formulas.put(cell, after);
        }
        for (var name : workbook.getAllNames()) {
            String before = name.getRefersToFormula();
            if (before != null) {
                String after = retarget(before, evaluation, FormulaType.NAMEDRANGE, name.getSheetIndex(), -1, sheet, active.getSheetName(), source, dr, dc);
                if (!before.equals(after)) names.put(name, after);
            }
        }
        formulas.forEach(Cell::setCellFormula); names.forEach(Name::setRefersToFormula);
        var values = new ArrayList<Value>();
        for (int r = 0; r < rows; r++) for (int c = 0; c < columns; c++) {
            var cell = book.cell(sheet, top + r, left + c, false);
            var type = cell == null ? CellType.BLANK : cell.getCellType();
            Object content = switch (type) {
                case STRING -> cell.getRichStringCellValue(); case FORMULA -> cell.getCellFormula();
                case NUMERIC -> cell.getNumericCellValue(); case BOOLEAN -> cell.getBooleanCellValue();
                case ERROR -> cell.getErrorCellValue(); default -> null;
            };
            values.add(new Value(type, content, cell == null ? workbook.getCellStyleAt(0) : cell.getCellStyle()));
        }
        for (int r = 0; r < rows; r++) for (int c = 0; c < columns; c++) {
            var cell = book.cell(sheet, top + r, left + c, false);
            if (cell != null) { cell.setBlank(); cell.setCellStyle(workbook.getCellStyleAt(0)); }
        }
        int index = 0;
        for (int r = 0; r < rows; r++) for (int c = 0; c < columns; c++) {
            var value = values.get(index++); var cell = book.cell(sheet, toRow + r, toColumn + c, true);
            cell.setBlank(); cell.setCellStyle(value.style());
            switch (value.type()) {
                case STRING -> cell.setCellValue((RichTextString) value.content());
                case FORMULA -> cell.setCellFormula((String) value.content());
                case NUMERIC -> cell.setCellValue((Double) value.content());
                case BOOLEAN -> cell.setCellValue((Boolean) value.content());
                case ERROR -> cell.setCellErrorValue((Byte) value.content()); default -> { }
            }
        }
        book.resetEvaluator();
    }

    private static String retarget(String formula, XSSFEvaluationWorkbook evaluation, FormulaType type, int context, int row,
                                   int sheet, String name, CellRangeAddress source, int dr, int dc) {
        var tokens = FormulaParser.parse(formula, evaluation, type, context, row);
        boolean changed = false;
        for (var token : tokens) {
            boolean applies;
            if (token instanceof Pxg3D ref) {
                if (ref.getExternalWorkbookNumber() >= 0) continue;
                if (ref.getLastSheetName() != null) { unsupported("여러 시트를 묶는 3D 참조"); return formula; }
                applies = name.equalsIgnoreCase(ref.getSheetName());
            } else if (token instanceof RefPtg || token instanceof AreaPtg) {
                if (context < 0) { unsupported("시트가 지정되지 않은 이름 정의"); return formula; }
                applies = context == sheet;
            } else if (token instanceof Ref3DPtg || token instanceof Area3DPtg) {
                unsupported("이전 형식의 외부 참조"); return formula;
            } else continue;
            if (!applies) continue;
            if (token instanceof RefPtgBase ref && source.isInRange(ref.getRow(), ref.getColumn())) {
                ref.setRow(ref.getRow() + dr); ref.setColumn(ref.getColumn() + dc); changed = true;
            } else if (token instanceof AreaPtgBase area) {
                var range = new CellRangeAddress(area.getFirstRow(), area.getLastRow(), area.getFirstColumn(), area.getLastColumn());
                if (!range.intersects(source)) continue;
                if (!source.isInRange(range.getFirstRow(), range.getFirstColumn()) || !source.isInRange(range.getLastRow(), range.getLastColumn()))
                    unsupported("잘라낼 셀과 일부만 겹치는 수식 범위");
                area.setFirstRow(area.getFirstRow() + dr); area.setLastRow(area.getLastRow() + dr);
                area.setFirstColumn(area.getFirstColumn() + dc); area.setLastColumn(area.getLastColumn() + dc); changed = true;
            }
        }
        return changed ? FormulaRenderer.toFormulaString(evaluation, tokens) : formula;
    }
    private static void unsupported(String detail) { throw new IllegalArgumentException(detail + "의 잘라내기는 아직 지원하지 않습니다. 원본은 유지됩니다."); }
}
