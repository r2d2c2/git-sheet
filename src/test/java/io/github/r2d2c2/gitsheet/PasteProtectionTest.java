package io.github.r2d2c2.gitsheet;

import org.junit.jupiter.api.Test;
import org.apache.poi.ss.util.CellRangeAddress;
import java.util.List;
import static org.junit.jupiter.api.Assertions.*;

class PasteProtectionTest {
    @Test void everyInternalModeRejectsMergedTargetBeforeValuesOrStylesChange() throws Exception {
        try (var book = new Book()) {
            book.set(0, 0, 0, "new"); book.style(0, 0, 0, "bold", "true");
            var copy = new CellClipboard(book, 0, 0, 0, 2, 1);
            book.set(0, 3, 0, "keep");
            book.workbook().getSheetAt(0).addMergedRegion(CellRangeAddress.valueOf("A5:B5"));
            int styles = book.workbook().getNumCellStyles();
            short style = book.cell(0, 3, 0, false).getCellStyle().getIndex();
            for (var mode : CellClipboard.Mode.values()) {
                assertThrows(IllegalArgumentException.class, () -> copy.paste(book, 0, 3, 0, mode));
                assertEquals("keep", book.raw(0, 3, 0));
                assertEquals(styles, book.workbook().getNumCellStyles());
                assertEquals(style, book.cell(0, 3, 0, false).getCellStyle().getIndex());
            }
            copy.paste(book, 0, 6, 0, CellClipboard.Mode.ALL);
            assertEquals("new", book.raw(0, 6, 0));
        }
    }
    @Test void internalPasteRejectsArrayAndProtectedTargetsInEveryMode() throws Exception {
        try (var book = new Book()) {
            book.set(0, 0, 0, "new");
            var copy = new CellClipboard(book, 0, 0, 0, 2, 1);
            book.set(0, 3, 0, "keep");
            var sheet = book.workbook().getSheetAt(0);
            sheet.setArrayFormula("ROW(A5:A6)", CellRangeAddress.valueOf("A5:A6"));
            for (var mode : CellClipboard.Mode.values()) {
                assertThrows(IllegalArgumentException.class, () -> copy.paste(book, 0, 3, 0, mode));
                assertEquals("keep", book.raw(0, 3, 0));
                assertTrue(book.cell(0, 4, 0, false).isPartOfArrayFormulaGroup());
            }
            sheet.protectSheet("test");
            for (var mode : CellClipboard.Mode.values())
                assertThrows(IllegalArgumentException.class, () -> copy.paste(book, 0, 8, 0, mode));
            assertNull(sheet.getRow(8));
        }
    }
    @Test void mergedCopyRejectsAnchorAndCoveredFragments() throws Exception {
        try (var book = new Book()) {
            book.workbook().getSheetAt(0).addMergedRegion(CellRangeAddress.valueOf("A1:B2"));
            assertThrows(IllegalArgumentException.class, () -> new CellClipboard(book, 0, 0, 0, 1, 1));
            assertThrows(IllegalArgumentException.class, () -> new CellClipboard(book, 0, 1, 1, 1, 1));
            assertThrows(IllegalArgumentException.class, () -> new CellClipboard(book, 0, 0, 0, 2, 2));
        }
    }
    @Test void externalPasteChecksAllRaggedRowsBeforeWriting() throws Exception {
        try (var book = new Book()) {
            book.set(0, 0, 0, "keep");
            var sheet = book.workbook().getSheetAt(0);
            sheet.addMergedRegion(CellRangeAddress.valueOf("B2:C2"));
            var bad = List.of(List.of("new"), List.of("1", "2"));
            assertThrows(IllegalArgumentException.class, () -> SheetEdits.pasteText(book, 0, 0, 0, bad, false));
            assertEquals("keep", book.raw(0, 0, 0)); assertNull(book.cell(0, 1, 0, false));
            // A ragged row does not write its missing B2 cell, so that merge is irrelevant.
            var good = List.of(List.of("=1+1", "001"), List.of("ok"));
            book.transaction(b -> SheetEdits.pasteText(b, 0, 0, 0, good, true));
            assertEquals("=1+1", book.raw(0, 0, 0)); assertEquals("001", book.raw(0, 0, 1));
            book.undo(); assertEquals("keep", book.raw(0, 0, 0));
        }
    }
    @Test void externalPasteRefusesArraysProtectionAndOverflowWithoutPartialWrite() throws Exception {
        try (var book = new Book()) {
            book.set(0, 0, 0, "keep");
            var sheet = book.workbook().getSheetAt(0);
            sheet.setArrayFormula("ROW(A2:A3)", CellRangeAddress.valueOf("A2:A3"));
            var values = List.of(List.of("first"), List.of("second"));
            assertThrows(IllegalArgumentException.class, () -> SheetEdits.pasteText(book, 0, 0, 0, values, false));
            assertEquals("keep", book.raw(0, 0, 0));
            assertThrows(IllegalArgumentException.class, () -> SheetEdits.pasteText(book, 0, Book.MAX_ROWS - 1, 0, values, false));
            assertNull(sheet.getRow(Book.MAX_ROWS - 1));
            sheet.protectSheet("test");
            assertThrows(IllegalArgumentException.class, () -> SheetEdits.pasteText(book, 0, 0, 0, List.of(List.of("new")), false));
            assertEquals("keep", book.raw(0, 0, 0));
        }
    }
}
