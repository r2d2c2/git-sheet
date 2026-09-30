package io.github.r2d2c2.gitsheet;

import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

/** Stable row sorting inside an explicit rectangle. Caller owns transaction/undo. */
public final class RangeSort {
    private RangeSort() {}
    public static final long MAX_CELLS = 1_000_000;
    private record Value(CellType type, Object content, CellStyle style) {}
    private record Key(int rank, Object value) {}
    private record SourceRow(int row, Key key) {}

    public static void sort(Book book, int sheetIndex, CellRangeAddress range, int keyColumn, boolean header, boolean descending) {
        int top = range.getFirstRow(), bottom = range.getLastRow(), left = range.getFirstColumn(), right = range.getLastColumn();
        book.cell(sheetIndex, top, left, false); book.cell(sheetIndex, bottom, right, false);
        if (top > bottom || left > right || keyColumn < left || keyColumn > right) throw new IllegalArgumentException("정렬 기준 열은 지정한 범위 안에 있어야 합니다.");
        if ((long) (bottom - top + 1) * (right - left + 1) > MAX_CELLS) throw new IllegalArgumentException("한 번에 정렬할 수 있는 범위는 100만 셀까지입니다.");
        int first = header ? top + 1 : top;
        if (first >= bottom) return;
        var sheet = book.workbook().getSheetAt(sheetIndex);
        if (sheet.getProtect()) unsupported("보호된 시트");
        for (var merged : sheet.getMergedRegions()) if (merged.intersects(range)) unsupported("병합 셀");
        if (!sheet.getTables().isEmpty() || !sheet.getDataValidations().isEmpty()
                || sheet.getSheetConditionalFormatting().getNumConditionalFormattings() > 0 || sheet.getDrawingPatriarch() != null)
            unsupported("표·유효성 검사·조건부 서식·그림이 있는 시트");
        for (int r = first; r <= bottom; r++) for (int c = left; c <= right; c++) {
            var cell = book.cell(sheetIndex, r, c, false);
            if (cell != null && (cell.isPartOfArrayFormulaGroup() || cell.getCellComment() != null || cell.getHyperlink() != null))
                unsupported("배열 수식·메모·하이퍼링크 셀");
        }
        var evaluator = book.workbook().getCreationHelper().createFormulaEvaluator();
        var order = new ArrayList<SourceRow>();
        for (int r = first; r <= bottom; r++) order.add(new SourceRow(r, key(book.cell(sheetIndex, r, keyColumn, false), evaluator)));
        order.sort((a, b) -> compare(a.key(), b.key(), descending));
        // Capture and translate before writing: overlapping rows never overwrite inputs.
        var values = new ArrayList<Value>();
        int targetRow = first;
        for (var source : order) {
            for (int c = left; c <= right; c++) {
                var cell = book.cell(sheetIndex, source.row(), c, false);
                var type = cell == null ? CellType.BLANK : cell.getCellType();
                Object content = switch (type) {
                    case STRING -> cell.getRichStringCellValue(); case NUMERIC -> cell.getNumericCellValue();
                    case BOOLEAN -> cell.getBooleanCellValue(); case ERROR -> cell.getErrorCellValue();
                    case FORMULA -> SheetEdits.translate(book, sheetIndex, cell.getCellFormula(), source.row(), c, targetRow, c);
                    default -> null;
                };
                values.add(new Value(type, content, cell == null ? book.workbook().getCellStyleAt(0) : cell.getCellStyle()));
            }
            targetRow++;
        }
        int index = 0;
        for (int r = first; r <= bottom; r++) for (int c = left; c <= right; c++) {
            var value = values.get(index++);
            var cell = book.cell(sheetIndex, r, c, value.type() != CellType.BLANK || value.style().getIndex() != 0);
            if (cell == null) continue;
            cell.setBlank(); cell.setCellStyle(value.style());
            switch (value.type()) {
                case STRING -> cell.setCellValue((RichTextString) value.content()); case NUMERIC -> cell.setCellValue((Double) value.content());
                case BOOLEAN -> cell.setCellValue((Boolean) value.content()); case ERROR -> cell.setCellErrorValue((Byte) value.content());
                case FORMULA -> cell.setCellFormula((String) value.content()); default -> { }
            }
        }
        book.resetEvaluator();
    }
    private static Key key(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) return new Key(4, null);
        var value = evaluator.evaluate(cell);
        if (value == null) return new Key(4, null);
        return switch (value.getCellType()) {
            case NUMERIC -> new Key(0, value.getNumberValue());
            case STRING -> value.getStringValue().isEmpty() ? new Key(4, null) : new Key(1, value.getStringValue());
            case BOOLEAN -> new Key(2, value.getBooleanValue()); case ERROR -> new Key(3, value.getErrorValue());
            default -> new Key(4, null);
        };
    }
    private static int compare(Key a, Key b, boolean descending) {
        if (a.rank() == 4 || b.rank() == 4) return Integer.compare(a.rank(), b.rank());
        int result = Integer.compare(a.rank(), b.rank());
        if (result == 0) result = switch (a.rank()) {
            case 0 -> Double.compare((Double) a.value(), (Double) b.value());
            case 1 -> ((String) a.value()).compareToIgnoreCase((String) b.value());
            case 2 -> Boolean.compare((Boolean) a.value(), (Boolean) b.value());
            default -> Byte.compare((Byte) a.value(), (Byte) b.value());
        };
        return descending ? -Integer.signum(result) : result;
    }
    private static void unsupported(String detail) { throw new IllegalArgumentException(detail + "의 데이터 정렬은 아직 지원하지 않습니다. 원본은 유지됩니다."); }
}
