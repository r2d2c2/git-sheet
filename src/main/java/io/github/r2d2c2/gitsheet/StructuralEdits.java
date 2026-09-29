package io.github.r2d2c2.gitsheet;

import java.util.*;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.helpers.BaseRowColShifter;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.helpers.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;

/** Whole-row/column edits. Call inside Book.transaction for rollback and one-step undo. */
public final class StructuralEdits {
    private StructuralEdits() {}
    public enum Axis { ROW, COLUMN }

    public static void change(Book book, int sheetIndex, Axis axis, int start, int count, boolean insert) {
        int limit = axis == Axis.ROW ? SpreadsheetVersion.EXCEL2007.getMaxRows() : SpreadsheetVersion.EXCEL2007.getMaxColumns();
        if (start < 0 || count < 1 || start >= limit || count > limit - start)
            throw new IllegalArgumentException("행·열 범위가 Excel 한도를 벗어납니다.");
        var sheet = book.workbook().getSheetAt(sheetIndex);
        // POI otherwise silently leaves unparseable formulas unchanged during a shift.
        var evaluation = XSSFEvaluationWorkbook.create(book.workbook());
        for (var s : book.workbook()) for (var row : s) for (var cell : row) {
            if (cell.isPartOfArrayFormulaGroup()) throw new IllegalArgumentException("배열 수식이 있는 문서의 행·열 구조 편집은 아직 지원하지 않습니다.");
            if (cell.getCellType() == CellType.FORMULA)
                FormulaParser.parse(cell.getCellFormula(), evaluation, FormulaType.CELL, book.workbook().getSheetIndex(s), row.getRowNum());
        }
        if (!sheet.getTables().isEmpty()) throw new IllegalArgumentException("Excel 표가 있는 시트의 구조 편집은 아직 지원하지 않습니다.");
        var merges = sheet.getMergedRegions();
        var columns = new ArrayList<CTCol>();
        for (var group : sheet.getCTWorksheet().getColsList()) for (var column : group.getColList()) columns.add((CTCol) column.copy());
        if (insert) {
            for (var row : sheet) {
                if (axis == Axis.ROW && row.getRowNum() >= limit - count) overflow();
                if (axis == Axis.COLUMN) for (var cell : row) if (cell.getColumnIndex() >= limit - count) overflow();
            }
            for (var merge : merges) if ((axis == Axis.ROW ? merge.getLastRow() : merge.getLastColumn()) >= limit - count) overflow();
            if (axis == Axis.COLUMN) for (var column : columns) if (column.getMax() > limit - count) overflow();
        }
        var pane = sheet.getPaneInformation();
        int freezeRows = pane != null && pane.isFreezePane() ? pane.getHorizontalSplitPosition() : 0;
        int freezeColumns = pane != null && pane.isFreezePane() ? pane.getVerticalSplitPosition() : 0;
        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) sheet.removeMergedRegion(i);

        if (!insert) {
            // Remove underlying XML too; POI column shifts can otherwise resurrect
            // overwritten cells when rebuilding row maps, particularly at the tail.
            var rows = new ArrayList<XSSFRow>(); sheet.forEach(r -> rows.add((XSSFRow) r));
            for (var row : rows) {
                if (axis == Axis.ROW && row.getRowNum() >= start && row.getRowNum() < start + count) {
                    for (var cell : row) { cell.removeCellComment(); cell.removeHyperlink(); }
                    sheet.removeRow(row);
                } else if (axis == Axis.COLUMN) {
                    var cells = new ArrayList<XSSFCell>(); row.forEach(c -> cells.add((XSSFCell) c));
                    for (var cell : cells) if (cell.getColumnIndex() >= start && cell.getColumnIndex() < start + count) {
                        cell.removeCellComment(); cell.removeHyperlink(); row.removeCell(cell);
                    }
                }
            }
        }

        int from = insert ? start : start + count;
        // Extend the empty source beyond the grid when needed: POI's overwrite
        // detection must cover every deleted coordinate, including a large tail cut.
        int end = insert ? limit - count - 1 : Math.max(limit - 1, from + count - 1);
        int delta = insert ? count : -count;
        if (insert) {
            // A virtual source beyond the grid makes POI invalidate references to the
            // deleted tail without ever creating an out-of-bounds cell.
            var formulaShift = axis == Axis.ROW
                    ? FormulaShifter.createForRowShift(sheetIndex, sheet.getSheetName(), limit, limit + count - 1, -count, SpreadsheetVersion.EXCEL2007)
                    : FormulaShifter.createForColumnShift(sheetIndex, sheet.getSheetName(), limit, limit + count - 1, -count, SpreadsheetVersion.EXCEL2007);
            BaseRowColShifter helper = axis == Axis.ROW ? new XSSFRowShifter(sheet) : new XSSFColumnShifter(sheet);
            helper.updateFormulas(formulaShift); helper.updateNamedRanges(formulaShift);
            helper.updateConditionalFormatting(formulaShift); helper.updateHyperlinks(formulaShift);
        }
        if (from <= end) {
            if (axis == Axis.ROW) sheet.shiftRows(from, end, delta, true, false);
            else sheet.shiftColumns(from, end, delta);
        }
        for (var merge : merges) {
            int[] interval = interval(axis == Axis.ROW ? merge.getFirstRow() : merge.getFirstColumn(), axis == Axis.ROW ? merge.getLastRow() : merge.getLastColumn(), start, count, insert);
            if (interval == null) continue;
            if (axis == Axis.ROW) { merge.setFirstRow(interval[0]); merge.setLastRow(interval[1]); }
            else { merge.setFirstColumn(interval[0]); merge.setLastColumn(interval[1]); }
            if (merge.getNumberOfCells() > 1) sheet.addMergedRegion(merge);
        }
        if (axis == Axis.COLUMN) {
            var xml = sheet.getCTWorksheet();
            while (xml.sizeOfColsArray() > 0) xml.removeCols(0);
            var group = xml.addNewCols();
            for (var column : columns) {
                // Inserted columns use the default layout; split an existing span at the gap.
                int first = (int) column.getMin() - 1, last = (int) column.getMax() - 1;
                if (insert && first < start && last >= start) {
                    var left = group.addNewCol(); left.set(column); left.setMax(start);
                    var right = group.addNewCol(); right.set(column); right.setMin(start + count + 1); right.setMax(last + count + 1);
                } else {
                    int[] range = interval(first, last, start, count, insert);
                    if (range != null) { var next = group.addNewCol(); next.set(column); next.setMin(range[0] + 1); next.setMax(range[1] + 1); }
                }
            }
        }
        if (freezeRows > 0 || freezeColumns > 0) {
            if (axis == Axis.ROW) freezeRows = boundary(freezeRows, start, count, insert);
            else freezeColumns = boundary(freezeColumns, start, count, insert);
            sheet.createFreezePane(freezeColumns, freezeRows);
        }
        book.resetEvaluator();
    }
    private static void overflow() { throw new IllegalArgumentException("마지막 행·열에 데이터나 서식이 있어 삽입할 수 없습니다. 이동할 공간을 먼저 확보하세요."); }
    private static int boundary(int boundary, int start, int count, boolean insert) {
        return boundary <= start ? boundary : insert ? boundary + count : Math.max(start, boundary - count);
    }
    private static int[] interval(int first, int last, int start, int count, boolean insert) {
        if (insert) return new int[] { first >= start ? first + count : first, last >= start ? last + count : last };
        int end = start + count - 1;
        if (first >= start && last <= end) return null;
        return new int[] { first > end ? first - count : Math.min(first, start), last > end ? last - count : Math.min(last, start - 1) };
    }
}
