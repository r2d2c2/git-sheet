package io.github.r2d2c2.gitsheet;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.apache.poi.ss.util.CellRangeAddress;
import java.nio.file.Path;
import static io.github.r2d2c2.gitsheet.StructuralEdits.Axis.*;
import static org.junit.jupiter.api.Assertions.*;

class StructuralEditsTest {
    @Test void largeTailInsertionInvalidatesOverflowReferences() throws Exception {
        for (var axis : StructuralEdits.Axis.values()) try (var b = new Book()) {
            int limit = axis == ROW ? 1_048_576 : 16_384;
            b.set(0, 0, 0, "=" + Book.address(axis == ROW ? limit - 3 : 1, axis == COLUMN ? limit - 3 : 1));
            b.set(0, axis == ROW ? limit - 10 : 1, axis == COLUMN ? limit - 10 : 1, "edge");
            StructuralEdits.change(b, 0, axis, limit - 10, 9, true);
            assertEquals("=#REF!", b.raw(0, 0, 0));
            assertEquals("edge", b.raw(0, axis == ROW ? limit - 1 : 1, axis == COLUMN ? limit - 1 : 1));
        }
    }
    @Test void wholeColumnAndWholeRowReferencesStayValid() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "=SUM(B:B)"); b.set(0, 0, 1, "=SUM(3:3)");
            StructuralEdits.change(b, 0, ROW, 1, 2, true);
            assertEquals("=SUM(B:B)", b.raw(0, 0, 0)); assertEquals("=SUM($A5:$XFD5)", b.raw(0, 0, 1));
            StructuralEdits.change(b, 0, COLUMN, 1, 2, true);
            assertEquals("=SUM(D:D)", b.raw(0, 0, 0)); assertEquals("=SUM($A5:$XFD5)", b.raw(0, 0, 3));
        }
    }
    @Test void deleteLargeRangeNearTailRemovesEveryCellAndReference() throws Exception {
        for (var axis : StructuralEdits.Axis.values()) try (var b = new Book()) {
            int limit = axis == ROW ? 1_048_576 : 16_384;
            for (int i = limit - 10; i < limit; i++) b.set(0, axis == ROW ? i : 0, axis == COLUMN ? i : 0, "v" + i);
            b.set(0, 0, 0, "=" + Book.address(axis == ROW ? limit - 3 : 0, axis == COLUMN ? limit - 3 : 0));
            StructuralEdits.change(b, 0, axis, limit - 10, 9, false);
            assertEquals("=#REF!", b.raw(0, 0, 0));
            assertEquals("v" + (limit - 1), b.raw(0, axis == ROW ? limit - 10 : 0, axis == COLUMN ? limit - 10 : 0));
            for (int i = limit - 9; i < limit; i++) assertEquals("", b.raw(0, axis == ROW ? i : 0, axis == COLUMN ? i : 0));
        }
    }
    @Test void insertRowsUpdatesCrossSheetAbsoluteRangeAndBlankReferencesWithUndo() throws Exception {
        try (var b = new Book()) {
            b.workbook().setSheetName(0, "Data"); b.workbook().createSheet("Summary");
            b.set(0, 1, 0, "5"); b.style(0, 1, 0, "bold", "true");
            b.workbook().getSheetAt(0).getRow(1).setHeightInPoints(32);
            b.set(1, 0, 0, "=SUM(Data!$A$1:$A$5)+Data!A100");
            b.transaction(x -> StructuralEdits.change(x, 0, ROW, 1, 2, true));
            assertEquals("5", b.raw(0, 3, 0)); assertEquals("", b.raw(0, 1, 0));
            assertEquals(32, b.workbook().getSheetAt(0).getRow(3).getHeightInPoints());
            assertTrue(b.workbook().getFontAt(b.cell(0, 3, 0, false).getCellStyle().getFontIndex()).getBold());
            assertEquals("=SUM(Data!$A$1:$A$7)+Data!A102", b.raw(1, 0, 0));
            b.undo(); assertEquals("5", b.raw(0, 1, 0));
            b.redo(); assertEquals("5", b.raw(0, 3, 0));
        }
    }
    @Test void deleteRowsInvalidatesRemovedReferencesAndShrinksRanges() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "=A2"); b.set(0, 0, 1, "=SUM(C1:C5)"); b.set(0, 3, 2, "7");
            StructuralEdits.change(b, 0, ROW, 1, 2, false);
            assertEquals("=#REF!", b.raw(0, 0, 0)); assertEquals("=SUM(C1:C3)", b.raw(0, 0, 1));
            assertEquals("7", b.raw(0, 1, 2)); assertEquals("7", b.display(0, 0, 1));
        }
    }
    @Test void columnInsertAndDeletePreserveLayoutAndAdjustFormulas() throws Exception {
        try (var b = new Book()) {
            var s = b.workbook().getSheetAt(0);
            s.setColumnWidth(1, 30 * 256); s.setColumnHidden(1, true);
            b.set(0, 0, 1, "9"); b.set(0, 1, 0, "=$B$1+Z100");
            StructuralEdits.change(b, 0, COLUMN, 1, 2, true);
            assertEquals("9", b.raw(0, 0, 3)); assertEquals("=$D$1+AB100", b.raw(0, 1, 0));
            assertEquals(30 * 256, s.getColumnWidth(3)); assertTrue(s.isColumnHidden(3)); assertFalse(s.isColumnHidden(1));
            StructuralEdits.change(b, 0, COLUMN, 2, 2, false);
            assertEquals("=#REF!+Z100", b.raw(0, 1, 0)); assertEquals("", b.raw(0, 0, 1));
        }
    }
    @Test void lastRowAndColumnDeletionInvalidateReferences() throws Exception {
        for (var axis : StructuralEdits.Axis.values()) try (var b = new Book()) {
            int last = axis == ROW ? 1_048_575 : 16_383;
            b.set(0, 0, 0, axis == ROW ? "=A1048576" : "=XFD1");
            b.set(0, axis == ROW ? last : 0, axis == COLUMN ? last : 0, "12");
            StructuralEdits.change(b, 0, axis, last, 1, false);
            assertEquals("=#REF!", b.raw(0, 0, 0));
            assertEquals("", b.raw(0, axis == ROW ? last : 0, axis == COLUMN ? last : 0));
        }
    }
    @Test void insertionRejectsDataLossAndHandlesEmptyGridTail() throws Exception {
        for (var axis : StructuralEdits.Axis.values()) try (var b = new Book()) {
            int last = axis == ROW ? 1_048_575 : 16_383;
            b.set(0, axis == ROW ? last : 0, axis == COLUMN ? last : 0, "precious");
            assertThrows(IllegalArgumentException.class, () -> b.transaction(x -> StructuralEdits.change(x, 0, axis, 1, 1, true)));
            assertEquals("precious", b.raw(0, axis == ROW ? last : 0, axis == COLUMN ? last : 0));
        }
        try (var b = new Book()) {
            b.set(0, 0, 0, "=A1048576+XFD2");
            StructuralEdits.change(b, 0, ROW, 1_048_575, 1, true);
            StructuralEdits.change(b, 0, COLUMN, 16_383, 1, true);
            assertEquals("=#REF!+#REF!", b.raw(0, 0, 0));
        }
    }
    @Test void mergesExpandShrinkAndDisappearWithDeletedRange() throws Exception {
        try (var b = new Book()) {
            var s = b.workbook().getSheetAt(0); s.addMergedRegion(CellRangeAddress.valueOf("A1:C4")); s.createFreezePane(2, 2);
            StructuralEdits.change(b, 0, ROW, 2, 2, true);
            assertEquals("A1:C6", s.getMergedRegion(0).formatAsString());
            StructuralEdits.change(b, 0, COLUMN, 1, 1, true);
            assertEquals("A1:D6", s.getMergedRegion(0).formatAsString()); assertEquals(3, s.getPaneInformation().getVerticalSplitPosition());
            StructuralEdits.change(b, 0, ROW, 0, 3, false);
            assertEquals("A1:D3", s.getMergedRegion(0).formatAsString());
            StructuralEdits.change(b, 0, COLUMN, 0, 4, false); assertEquals(0, s.getNumMergedRegions());
        }
    }
    @Test void namedRangesMoveAndFilesRoundTrip(@TempDir Path dir) throws Exception {
        try (var b = new Book()) {
            var s = b.workbook().getSheetAt(0); b.workbook().setSheetName(0, "Data");
            b.set(0, 1, 1, "42"); s.setColumnWidth(120, 40 * 256); s.setColumnHidden(120, true);
            var name = b.workbook().createName(); name.setNameName("Total"); name.setRefersToFormula("Data!$B$2");
            StructuralEdits.change(b, 0, COLUMN, 1, 2, true);
            assertEquals("Data!$D$2", name.getRefersToFormula());
            var xlsx = dir.resolve("book.xlsx"); BookFiles.exportExcel(b, xlsx);
            try (var loaded = BookFiles.importExcel(xlsx)) { assertEquals("42", loaded.raw(0, 1, 3)); assertEquals("Data!$D$2", loaded.workbook().getName("Total").getRefersToFormula()); }
            var text = dir.resolve("book.gsheet"); BookFiles.write(b, text);
            try (var loaded = BookFiles.read(text)) {
                assertEquals("42", loaded.raw(0, 1, 3)); assertEquals(40 * 256, loaded.workbook().getSheetAt(0).getColumnWidth(122));
                assertTrue(loaded.workbook().getSheetAt(0).isColumnHidden(122));
            }
        }
    }
}
