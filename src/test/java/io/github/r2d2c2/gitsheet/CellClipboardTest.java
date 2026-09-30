package io.github.r2d2c2.gitsheet;

import org.junit.jupiter.api.Test;
import org.apache.poi.ss.usermodel.*;
import static org.junit.jupiter.api.Assertions.*;

class CellClipboardTest {
    @Test void copyDoesNotSnapshotUnselectedWorkbookAndSurvivesSourceClose() throws Exception {
        CellClipboard copy;
        try (var source = new Book(new org.apache.poi.xssf.usermodel.XSSFWorkbook() {
            @Override protected void commit() { throw new AssertionError("Clipboard must not serialize the source workbook"); }
        })) {
            source.workbook().createSheet("Source");
            source.set(0, 0, 0, "selected"); source.style(0, 0, 0, "bold", "true");
            for (int i = 1; i <= 2000; i++) source.set(0, i, 10, "unselected-" + i + "x".repeat(500));
            copy = new CellClipboard(source, 0, 0, 0, 1, 1);
        }
        try (var target = new Book()) {
            copy.paste(target, 0, 2, 2, CellClipboard.Mode.ALL);
            assertEquals("selected", target.raw(0, 2, 2));
            assertTrue(target.workbook().getFontAt(target.cell(0, 2, 2, false).getCellStyle().getFontIndex()).getBold());
        }
    }
    @Test void namesAndCrossSheetDependenciesAreFrozenWithoutRetainingWorkbookData() throws Exception {
        try (var b = new Book()) {
            b.workbook().createSheet("Source Data"); b.set(1, 0, 0, "5");
            var unused = b.workbook().createName(); unused.setNameName("Unused"); unused.setRefersToFormula("'Source Data'!$A$1");
            var name = b.workbook().createName(); name.setNameName("Amount"); name.setRefersToFormula("'Source Data'!$A$1");
            b.set(0, 0, 0, "=Amount+'Source Data'!A1");
            var copy = new CellClipboard(b, 0, 0, 0, 1, 1);
            b.workbook().removeName(unused); b.set(1, 0, 0, "20");
            copy.paste(b, 0, 1, 0, CellClipboard.Mode.VALUES); assertEquals("10", b.raw(0, 1, 0));
            copy.paste(b, 0, 2, 0, CellClipboard.Mode.ALL); assertEquals("=Amount+'Source Data'!A3", b.raw(0, 2, 0));
            copy.paste(b, 0, 3, 0, CellClipboard.Mode.ALL); assertEquals("=Amount+'Source Data'!A4", b.raw(0, 3, 0));
        }
    }
    @Test void formulaErrorsAndUnsupportedCalculationHaveDistinctPasteBehavior() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "=1/0");
            new CellClipboard(b, 0, 0, 0, 1, 1).paste(b, 0, 1, 0, CellClipboard.Mode.VALUES);
            assertEquals(CellType.ERROR, b.cell(0, 1, 0, false).getCellType()); assertEquals("#DIV/0!", b.raw(0, 1, 0));
            b.set(0, 0, 1, "=INFO(\"system\")"); b.style(0, 0, 1, "bold", "true");
            var copy = new CellClipboard(b, 0, 0, 1, 1, 1);
            copy.paste(b, 0, 1, 1, CellClipboard.Mode.ALL); assertEquals("=INFO(\"system\")", b.raw(0, 1, 1));
            b.set(0, 2, 1, "keep"); copy.paste(b, 0, 2, 1, CellClipboard.Mode.FORMATS);
            assertThrows(IllegalArgumentException.class, () -> b.transaction(x -> copy.paste(x, 0, 2, 1, CellClipboard.Mode.VALUES)));
            assertEquals("keep", b.raw(0, 2, 1));
        }
    }
    @Test void repeatedPasteDoesNotMutateStoredFormulaTokensAtGridBoundary() throws Exception {
        try (var b = new Book()) {
            b.set(0, 1, 1, "=A1+$A$1");
            var copy = new CellClipboard(b, 0, 1, 1, 1, 1);
            copy.paste(b, 0, 0, 0, CellClipboard.Mode.ALL); assertEquals("=#REF!+$A$1", b.raw(0, 0, 0));
            copy.paste(b, 0, 3, 3, CellClipboard.Mode.ALL); assertEquals("=C3+$A$1", b.raw(0, 3, 3));
        }
    }
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
