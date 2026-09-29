package io.github.r2d2c2.gitsheet;

import org.junit.jupiter.api.Test;
import org.apache.poi.ss.usermodel.*;
import static org.junit.jupiter.api.Assertions.*;

class SheetEditsTest {
    @Test void fillDownMovesRelativeReferencesAndPreservesAbsoluteReferences() throws Exception {
        try (var book = new Book()) {
            book.set(0, 0, 0, "10"); book.set(0, 1, 0, "20"); book.set(0, 2, 0, "30");
            book.set(0, 0, 1, "=A1+$A$1+A$1+$A1"); book.style(0, 0, 1, "bold", "true");
            book.transaction(b -> SheetEdits.fill(b, 0, 0, 2, 1, 1, SheetEdits.Direction.DOWN));
            assertEquals("=A3+$A$1+A$1+$A3", book.raw(0, 2, 1)); assertEquals("80", book.display(0, 2, 1));
            assertEquals(book.cell(0, 0, 1, false).getCellStyle().getIndex(), book.cell(0, 2, 1, false).getCellStyle().getIndex());
            book.undo(); assertEquals("", book.raw(0, 2, 1));
        }
    }
    @Test void fillRightAdjustsCrossSheetAndRangeReferences() throws Exception {
        try (var book = new Book()) {
            book.workbook().createSheet("Data"); book.set(0, 0, 0, "=SUM(Data!A1:A3)+$B$1");
            SheetEdits.fill(book, 0, 0, 0, 0, 2, SheetEdits.Direction.RIGHT);
            assertEquals("=SUM(Data!C1:C3)+$B$1", book.raw(0, 0, 2));
        }
    }
    @Test void fillPreservesLiteralTypesAndBlankSourcesClearTargets() throws Exception {
        try (var book = new Book()) {
            book.set(0, 0, 0, "'=not-a-formula"); book.set(0, 0, 1, "TRUE"); book.set(0, 1, 2, "99");
            SheetEdits.fill(book, 0, 0, 1, 0, 2, SheetEdits.Direction.DOWN);
            assertEquals(CellType.STRING, book.cell(0, 1, 0, false).getCellType()); assertEquals("=not-a-formula", book.raw(0, 1, 0));
            assertEquals(CellType.BOOLEAN, book.cell(0, 1, 1, false).getCellType()); assertEquals("", book.raw(0, 1, 2));
        }
    }
    @Test void replaceAllIsLiteralAndCanExcludeFormulas() throws Exception {
        try (var book = new Book()) {
            book.set(0, 0, 0, "A.b A.B"); book.set(0, 0, 1, "=A1");
            assertEquals(1, SheetEdits.replaceAll(book, 0, "a.b", "$5\\value", false, false));
            assertEquals("$5\\value $5\\value", book.raw(0, 0, 0)); assertEquals("=A1", book.raw(0, 0, 1));
            assertEquals(1, SheetEdits.replaceAll(book, 0, "A1", "A2", true, true)); assertEquals("=A2", book.raw(0, 0, 1));
        }
    }
    @Test void failedFormulaReplaceRollsBackAllCells() throws Exception {
        try (var book = new Book()) {
            book.set(0, 0, 0, "A1"); book.set(0, 0, 1, "=A1");
            assertThrows(RuntimeException.class, () -> book.transaction(b -> SheetEdits.replaceAll(b, 0, "A1", "SUM(", true, true)));
            assertEquals("A1", book.raw(0, 0, 0)); assertEquals("=A1", book.raw(0, 0, 1));
        }
    }
}
