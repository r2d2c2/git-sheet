package io.github.r2d2c2.gitsheet;

import java.util.List;
import java.nio.file.*;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.apache.poi.ss.util.CellRangeAddress;
import static io.github.r2d2c2.gitsheet.RangeFilter.Operation.*;
import static org.junit.jupiter.api.Assertions.*;

class RangeFilterTest {
    private static RangeFilter.Spec spec(String range, boolean header, RangeFilter.Condition... conditions) {
        return new RangeFilter.Spec(CellRangeAddress.valueOf(range), header, List.of(conditions));
    }
    private static RangeFilter.Condition condition(int column, RangeFilter.Operation operation, String value) {
        return new RangeFilter.Condition(column, operation, value);
    }
    @Test void combinesTextAndNumericConditionsAndKeepsOriginalCoordinates(@TempDir Path dir) throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "Header"); b.set(0, 1, 0, "Apple"); b.set(0, 1, 1, "10");
            b.set(0, 2, 0, "apple juice"); b.set(0, 2, 1, "=2+3"); b.set(0, 3, 0, "orange"); b.set(0, 3, 1, "12");
            var before = dir.resolve("before.gsheet"); var after = dir.resolve("after.gsheet"); BookFiles.write(b, before);
            assertEquals(List.of(1), RangeFilter.matchingRows(b, 0, spec("A1:B4", true, condition(0, CONTAINS, "APPLE"), condition(1, GREATER, "5"))));
            BookFiles.write(b, after); assertEquals(Files.readString(before), Files.readString(after));
            assertEquals("=2+3", b.raw(0, 2, 1));
        }
    }
    @Test void blankAndNumericTypesAreNotConfusedWithZeroOrNumericStrings() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "0"); b.set(0, 1, 0, "'0"); b.set(0, 2, 0, "=\"\""); b.set(0, 4, 0, " ");
            assertEquals(List.of(2, 3), RangeFilter.matchingRows(b, 0, spec("A1:A5", false, condition(0, BLANK, ""))));
            assertEquals(List.of(0), RangeFilter.matchingRows(b, 0, spec("A1:A5", false, condition(0, EQUAL, "0"))));
            assertEquals(List.of(0, 1, 4), RangeFilter.matchingRows(b, 0, spec("A1:A5", false, condition(0, NOT_BLANK, ""))));
        }
    }
    @Test void wholeRangeResultsExtendBeyondWindowAndRecomputeAfterEdit() throws Exception {
        try (var b = new Book()) {
            for (int r = 0; r < 2005; r++) b.set(0, r, 0, "1");
            var filter = spec("A1:A2005", false, condition(0, AT_LEAST, "1"));
            var result = RangeFilter.matchingRows(b, 0, filter); assertEquals(2005, result.size()); assertEquals(2004, result.getLast());
            b.transaction(x -> x.set(0, 1500, 0, "0")); assertFalse(RangeFilter.matchingRows(b, 0, filter).contains(1500));
            b.undo(); assertEquals(2005, RangeFilter.matchingRows(b, 0, filter).size());
        }
    }
    @Test void validationAndDefensiveCopiesProtectFilterDefinition() {
        assertThrows(IllegalArgumentException.class, () -> condition(0, GREATER, "NaN"));
        assertThrows(IllegalArgumentException.class, () -> spec("A1:A2", false, condition(1, BLANK, "")));
        assertThrows(IllegalArgumentException.class, () -> spec("A1:A100001", false, condition(0, BLANK, "")));
        var area = CellRangeAddress.valueOf("A1:A2"); var filter = new RangeFilter.Spec(area, false, List.of(condition(0, BLANK, "")));
        area.setLastRow(999); filter.range().setLastRow(999); assertEquals(1, filter.range().getLastRow());
    }
    @Test void textEqualityAndFormulaFailuresAreExplicit() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "TRUE"); b.set(0, 1, 0, "=1/0"); b.set(0, 2, 0, "True story");
            assertEquals(List.of(0), RangeFilter.matchingRows(b, 0, spec("A1:A3", false, condition(0, TEXT_EQUALS, "true"))));
            assertEquals(List.of(1), RangeFilter.matchingRows(b, 0, spec("A1:A3", false, condition(0, TEXT_EQUALS, "#DIV/0!"))));
            b.set(0, 0, 0, "=INFO(\"system\")");
            assertThrows(RuntimeException.class, () -> RangeFilter.matchingRows(b, 0, spec("A1:A3", false, condition(0, LESS, "2"))));
            assertEquals("=INFO(\"system\")", b.raw(0, 0, 0));
        }
    }
}
