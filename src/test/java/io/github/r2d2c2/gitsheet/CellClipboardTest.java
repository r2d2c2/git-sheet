package io.github.r2d2c2.gitsheet;

import org.junit.jupiter.api.Test;
import org.apache.poi.ss.usermodel.*;
import static org.junit.jupiter.api.Assertions.*;

class CellClipboardTest {
    @Test void copiesTypesAndStyleFromSnapshotAcrossUndo() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "'=SUM(A1)"); b.set(0, 0, 1, "TRUE"); b.set(0, 0, 2, "12");
            b.style(0, 0, 0, "bold", "true");
            var copy = new CellClipboard(b, 0, 0, 0, 1, 3);
            b.transaction(x -> x.set(0, 0, 0, "changed")); b.undo();
            b.transaction(x -> copy.paste(x, 0, 1, 0, CellClipboard.Mode.ALL));
            assertEquals(CellType.STRING, b.cell(0, 1, 0, false).getCellType());
            assertEquals("=SUM(A1)", b.raw(0, 1, 0));
            assertEquals(CellType.BOOLEAN, b.cell(0, 1, 1, false).getCellType());
            assertEquals(CellType.NUMERIC, b.cell(0, 1, 2, false).getCellType());
            assertTrue(b.workbook().getFontAt(b.cell(0, 1, 0, false).getCellStyle().getFontIndex()).getBold());
            b.undo(); assertEquals("", b.raw(0, 1, 0)); b.redo(); assertEquals("12", b.raw(0, 1, 2));
        }
    }
    @Test void formulasShiftRelativeReferencesAcrossSheets() throws Exception {
        try (var b = new Book()) {
            b.workbook().createSheet("Other");
            b.set(0, 0, 1, "=A1+$A$1+A$1+$A1");
            var copy = new CellClipboard(b, 0, 0, 1, 1, 1);
            copy.paste(b, 1, 2, 3, CellClipboard.Mode.ALL);
            assertEquals("=C3+$A$1+C$1+$A3", b.raw(1, 2, 3));
        }
    }
    @Test void valuesPasteFreezesCalculatedResultAndPreservesDestinationStyle() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "3"); b.set(0, 0, 1, "=A1*2"); b.style(0, 1, 1, "bold", "true");
            var copy = new CellClipboard(b, 0, 0, 1, 1, 1);
            b.set(0, 0, 0, "10");
            copy.paste(b, 0, 1, 1, CellClipboard.Mode.VALUES);
            assertEquals("6", b.raw(0, 1, 1)); assertEquals(CellType.NUMERIC, b.cell(0, 1, 1, false).getCellType());
            assertTrue(b.workbook().getFontAt(b.cell(0, 1, 1, false).getCellStyle().getFontIndex()).getBold());
        }
    }
    @Test void formatsOnlyPreservesContentAndOverlappingPasteUsesOriginalCells() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "first"); b.set(0, 0, 1, "second"); b.style(0, 0, 0, "bold", "true");
            var copy = new CellClipboard(b, 0, 0, 0, 1, 2);
            copy.paste(b, 0, 0, 1, CellClipboard.Mode.FORMATS); assertEquals("second", b.raw(0, 0, 1));
            copy.paste(b, 0, 0, 1, CellClipboard.Mode.ALL);
            assertEquals("first", b.raw(0, 0, 1)); assertEquals("second", b.raw(0, 0, 2));
        }
    }
    @Test void boundaryFailureDoesNotPartiallyPaste() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "original");
            var copy = new CellClipboard(b, 0, 0, 0, 2, 2);
            assertThrows(IllegalArgumentException.class, () -> b.transaction(x -> copy.paste(x, 0, Book.MAX_ROWS - 1, 0, CellClipboard.Mode.ALL)));
            assertEquals("original", b.raw(0, 0, 0)); assertEquals("", b.raw(0, Book.MAX_ROWS - 1, 0));
        }
    }
    @Test void editingLiteralStringsNeverChangesTheirTypeOrContent() throws Exception {
        try (var b = new Book()) {
            for (String literal : new String[] { "=A1", "TRUE", "12", "0012", "'text", "", "hello" }) {
                b.cell(0, 0, 0, true).setCellValue(literal);
                b.set(0, 0, 0, b.input(0, 0, 0));
                assertEquals(CellType.STRING, b.cell(0, 0, 0, false).getCellType()); assertEquals(literal, b.raw(0, 0, 0));
            }
        }
    }
}
