package io.github.r2d2c2.gitsheet;

import java.nio.file.Path;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import static org.junit.jupiter.api.Assertions.*;

class RangeSortTest {
    @Test void headerAndRowPayloadStayTogetherAndUndoRestoresOrder() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "Key"); b.set(0, 0, 1, "Value");
            b.set(0, 1, 0, "10"); b.set(0, 1, 1, "'=literal"); b.style(0, 1, 1, "bold", "true");
            b.set(0, 2, 0, "2"); b.set(0, 2, 1, "TRUE"); b.set(0, 3, 0, "2"); b.set(0, 3, 1, "second two");
            b.transaction(x -> RangeSort.sort(x, 0, CellRangeAddress.valueOf("A1:B4"), 0, true, false));
            assertEquals("Key", b.raw(0, 0, 0)); assertEquals("2", b.raw(0, 1, 0)); assertEquals("second two", b.raw(0, 2, 1));
            assertEquals(CellType.BOOLEAN, b.cell(0, 1, 1, false).getCellType());
            assertEquals(CellType.STRING, b.cell(0, 3, 1, false).getCellType()); assertEquals("=literal", b.raw(0, 3, 1));
            assertTrue(b.workbook().getFontAt(b.cell(0, 3, 1, false).getCellStyle().getFontIndex()).getBold());
            b.undo(); assertEquals("10", b.raw(0, 1, 0)); b.redo(); assertEquals("10", b.raw(0, 3, 0));
        }
    }
    @Test void mixedTypesAndBlanksHaveDeterministicOrderInBothDirections() throws Exception {
        try (var b = new Book()) {
            String[] values = { "", "b", "TRUE", "10", "a", "2", "A" };
            for (int i = 0; i < values.length; i++) { b.set(0, i, 0, values[i]); b.set(0, i, 1, "row" + i); }
            RangeSort.sort(b, 0, CellRangeAddress.valueOf("A1:B7"), 0, false, false);
            assertEquals("2", b.raw(0, 0, 0)); assertEquals("10", b.raw(0, 1, 0)); assertEquals("a", b.raw(0, 2, 0));
            assertEquals("A", b.raw(0, 3, 0)); assertEquals("", b.raw(0, 6, 0)); assertEquals("row0", b.raw(0, 6, 1));
            RangeSort.sort(b, 0, CellRangeAddress.valueOf("A1:B7"), 0, false, true);
            assertEquals("TRUE", b.raw(0, 0, 0)); assertEquals("b", b.raw(0, 1, 0)); assertEquals("", b.raw(0, 6, 0));
        }
    }
    @Test void formulasSortByCalculatedKeysAndRelativeReferencesMoveWithRows() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "10"); b.set(0, 1, 0, "2"); b.set(0, 0, 1, "=A1+$D$2"); b.set(0, 1, 1, "=A2+$D$2");
            b.set(0, 1, 3, "5"); b.set(0, 0, 3, "=A1");
            RangeSort.sort(b, 0, CellRangeAddress.valueOf("A1:B2"), 1, false, false);
            assertEquals("2", b.raw(0, 0, 0)); assertEquals("=A1+$D$2", b.raw(0, 0, 1)); assertEquals("7", b.display(0, 0, 1));
            assertEquals("15", b.display(0, 1, 1)); assertEquals("=A1", b.raw(0, 0, 3));
        }
    }
    @Test void sortsBeyondVisibleWindowAndSurvivesFileRoundTrips(@TempDir Path dir) throws Exception {
        try (var b = new Book()) {
            for (int r = 0; r < 1500; r++) b.set(0, r, 0, Integer.toString(1500 - r));
            RangeSort.sort(b, 0, CellRangeAddress.valueOf("A1:A1500"), 0, false, false);
            assertEquals("1", b.raw(0, 0, 0)); assertEquals("1001", b.raw(0, 1000, 0)); assertEquals("1500", b.raw(0, 1499, 0));
            var text = dir.resolve("sorted.gsheet"); BookFiles.write(b, text);
            try (var loaded = BookFiles.read(text)) { assertEquals("1001", loaded.raw(0, 1000, 0)); }
            var excel = dir.resolve("sorted.xlsx"); BookFiles.exportExcel(b, excel);
            try (var loaded = BookFiles.importExcel(excel)) { assertEquals("1500", loaded.raw(0, 1499, 0)); }
        }
    }
    @Test void rejectedSortLeavesValuesIntact() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "9"); b.set(0, 1, 0, "1");
            assertThrows(IllegalArgumentException.class, () -> b.transaction(x -> RangeSort.sort(x, 0, CellRangeAddress.valueOf("A1:A2"), 1, false, false)));
            assertThrows(IllegalArgumentException.class, () -> b.transaction(x -> RangeSort.sort(x, 0, CellRangeAddress.valueOf("A1:B1000000"), 0, false, false)));
            b.workbook().getSheetAt(0).addMergedRegion(CellRangeAddress.valueOf("A1:B1"));
            assertThrows(IllegalArgumentException.class, () -> b.transaction(x -> RangeSort.sort(x, 0, CellRangeAddress.valueOf("A1:B2"), 0, false, false)));
            assertEquals("9", b.raw(0, 0, 0)); assertEquals("1", b.raw(0, 1, 0));
        }
    }
    @Test void unsupportedFormulaKeyRollsBackAndErrorKeysArePreserved() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "=INFO(\"system\")"); b.set(0, 1, 0, "1");
            assertThrows(RuntimeException.class, () -> b.transaction(x -> RangeSort.sort(x, 0, CellRangeAddress.valueOf("A1:A2"), 0, false, false)));
            assertEquals("=INFO(\"system\")", b.raw(0, 0, 0));
            b.set(0, 0, 0, "=1/0"); RangeSort.sort(b, 0, CellRangeAddress.valueOf("A1:A2"), 0, false, false);
            assertEquals("1", b.raw(0, 0, 0)); assertEquals("#DIV/0!", b.display(0, 1, 0));
        }
    }
}
