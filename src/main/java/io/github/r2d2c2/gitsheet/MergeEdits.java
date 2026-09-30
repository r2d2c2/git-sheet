package io.github.r2d2c2.gitsheet;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Merge edits never discard covered values. Caller wraps mutations in Book.transaction. */
public final class MergeEdits {
    private MergeEdits() {}
    public static CellRangeAddress containing(XSSFSheet sheet, int row, int column) {
        for (var range : sheet.getMergedRegions()) if (range.isInRange(row, column)) return range;
        return null;
    }
    public static void merge(Book book, int index, CellRangeAddress range) {
        range.validate(org.apache.poi.ss.SpreadsheetVersion.EXCEL2007);
        if (range.getNumberOfCells() < 2) throw new IllegalArgumentException("병합할 셀을 두 개 이상 선택하세요.");
        var sheet = book.workbook().getSheetAt(index);
        if (sheet.getProtect()) throw new IllegalArgumentException("보호된 시트에서는 병합할 수 없습니다.");
        for (var table : sheet.getTables()) {
            var tableRange = new CellRangeAddress(table.getStartRowIndex(), table.getEndRowIndex(), table.getStartColIndex(), table.getEndColIndex());
            if (tableRange.intersects(range)) throw new IllegalArgumentException("Excel 표와 겹치는 셀은 병합할 수 없습니다.");
        }
        for (var existing : sheet.getMergedRegions()) {
            if (existing.equals(range)) return;
            if (existing.intersects(range)) throw new IllegalArgumentException("기존 병합을 해제한 뒤 다시 병합하세요.");
        }
        for (var row : sheet) for (var cell : row) {
            if (!range.isInRange(cell)) continue;
            if (cell.isPartOfArrayFormulaGroup()) throw new IllegalArgumentException("배열 수식은 병합할 수 없습니다.");
            boolean anchor = cell.getRowIndex() == range.getFirstRow() && cell.getColumnIndex() == range.getFirstColumn();
            if (!anchor && (cell.getCellType() != CellType.BLANK || cell.getCellComment() != null || cell.getHyperlink() != null))
                throw new IllegalArgumentException("좌상단 이외 셀에 값·메모·링크가 있습니다. 내용을 먼저 옮겨 주세요.");
        }
        sheet.addMergedRegion(range.copy());
    }
    public static void unmerge(Book book, int index, CellRangeAddress selection) {
        var sheet = book.workbook().getSheetAt(index);
        if (sheet.getProtect()) throw new IllegalArgumentException("보호된 시트에서는 병합을 해제할 수 없습니다.");
        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--)
            if (sheet.getMergedRegion(i).intersects(selection)) sheet.removeMergedRegion(i);
    }
}
