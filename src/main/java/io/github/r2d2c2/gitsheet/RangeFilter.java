package io.github.r2d2c2.gitsheet;

import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.NumberToTextConverter;

/** Read-only, AND-combined predicates evaluated against original sheet coordinates. */
public final class RangeFilter {
    private RangeFilter() {}
    public enum Operation {
        CONTAINS("문자 포함"), TEXT_EQUALS("문자 일치"), BLANK("빈 값"), NOT_BLANK("빈 값 아님"),
        EQUAL("숫자 ="), GREATER("숫자 >"), AT_LEAST("숫자 ≥"), LESS("숫자 <"), AT_MOST("숫자 ≤");
        private final String label;
        Operation(String label) { this.label = label; }
        @Override public String toString() { return label; }
        boolean numeric() { return ordinal() >= EQUAL.ordinal(); }
    }
    public record Condition(int column, Operation operation, String operand) {
        public Condition {
            Objects.requireNonNull(operation); operand = Objects.requireNonNull(operand);
            if (operation.numeric()) {
                try { if (!Double.isFinite(Double.parseDouble(operand.trim()))) throw new NumberFormatException(); }
                catch (NumberFormatException e) { throw new IllegalArgumentException("숫자 조건에는 유한한 숫자를 입력하세요."); }
            }
        }
    }
    public record Spec(CellRangeAddress range, boolean header, List<Condition> conditions) {
        public Spec {
            range = range.copy(); conditions = List.copyOf(conditions);
            if (range.getFirstRow() < 0 || range.getLastRow() >= Book.MAX_ROWS || range.getFirstColumn() < 0 || range.getLastColumn() >= Book.MAX_COLUMNS)
                throw new IllegalArgumentException("필터 범위가 Excel 셀 범위를 벗어납니다.");
            if (range.getLastRow() - range.getFirstRow() + 1 > 100_000) throw new IllegalArgumentException("필터는 한 번에 10만 행까지 검사합니다.");
            if (conditions.isEmpty()) throw new IllegalArgumentException("필터 조건을 하나 이상 선택하세요.");
            for (var condition : conditions) if (condition.column() < range.getFirstColumn() || condition.column() > range.getLastColumn())
                throw new IllegalArgumentException("조건 열은 필터 범위 안에 있어야 합니다.");
        }
        @Override public CellRangeAddress range() { return range.copy(); }
    }
    public static List<Integer> matchingRows(Book book, int sheet, Spec spec) {
        var evaluator = book.workbook().getCreationHelper().createFormulaEvaluator();
        var result = new ArrayList<Integer>(); var range = spec.range();
        var numeric = spec.conditions().stream().map(c -> c.operation().numeric() ? Double.parseDouble(c.operand().trim()) : 0.0).toList();
        for (int r = range.getFirstRow() + (spec.header() ? 1 : 0); r <= range.getLastRow(); r++) {
            boolean match = true;
            for (int i = 0; i < spec.conditions().size(); i++) {
                var condition = spec.conditions().get(i); var cell = book.cell(sheet, r, condition.column(), false);
                var value = cell == null ? null : evaluator.evaluate(cell);
                if (!matches(value, condition, numeric.get(i))) { match = false; break; }
            }
            if (match) result.add(r);
        }
        return List.copyOf(result);
    }
    private static boolean matches(CellValue value, Condition condition, double expected) {
        boolean blank = value == null || value.getCellType() == CellType.BLANK
                || (value.getCellType() == CellType.STRING && value.getStringValue().isEmpty());
        if (condition.operation() == Operation.BLANK) return blank;
        if (condition.operation() == Operation.NOT_BLANK) return !blank;
        if (condition.operation().numeric()) {
            if (blank || value.getCellType() != CellType.NUMERIC) return false;
            double number = value.getNumberValue();
            return switch (condition.operation()) {
                case EQUAL -> number == expected; case GREATER -> number > expected; case AT_LEAST -> number >= expected;
                case LESS -> number < expected; case AT_MOST -> number <= expected; default -> false;
            };
        }
        String text = blank ? "" : switch (value.getCellType()) {
            case STRING -> value.getStringValue(); case NUMERIC -> NumberToTextConverter.toText(value.getNumberValue());
            case BOOLEAN -> Boolean.toString(value.getBooleanValue()); case ERROR -> FormulaError.forInt(value.getErrorValue()).getString();
            default -> "";
        };
        text = text.toLowerCase(Locale.ROOT); String operand = condition.operand().toLowerCase(Locale.ROOT);
        return condition.operation() == Operation.CONTAINS ? text.contains(operand) : text.equals(operand);
    }
}
