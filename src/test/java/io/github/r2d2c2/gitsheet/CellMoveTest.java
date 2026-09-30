package io.github.r2d2c2.gitsheet;

import org.junit.jupiter.api.Test;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import static org.junit.jupiter.api.Assertions.*;

class CellMoveTest {
    @Test void moveRetargetsAbsoluteAndCrossSheetReferencesAndNamedRangesWithUndo() throws Exception {
        try (var b = new Book()) {
            b.workbook().setSheetName(0, "Data"); b.workbook().createSheet("Summary");
            b.set(0, 0, 0, "7"); b.style(0, 0, 0, "bold", "true");
            b.set(0, 3, 3, "=$A$1+A1"); b.set(1, 0, 0, "=Data!$A$1");
            var name = b.workbook().createName(); name.setNameName("Source"); name.setRefersToFormula("Data!$A$1");
            b.transaction(x -> CellMove.move(x, 0, 0, 0, 1, 1, 1, 1));
            assertEquals("", b.raw(0, 0, 0)); assertEquals("7", b.raw(0, 1, 1));
            assertEquals("=$B$2+B2", b.raw(0, 3, 3)); assertEquals("=Data!$B$2", b.raw(1, 0, 0));
            assertEquals("Data!$B$2", b.workbook().getName("Source").getRefersToFormula()); assertEquals("14", b.display(0, 3, 3));
            assertTrue(b.workbook().getFontAt(b.cell(0, 1, 1, false).getCellStyle().getFontIndex()).getBold());
            b.undo(); assertEquals("7", b.raw(0, 0, 0)); assertEquals("=Data!$A$1", b.raw(1, 0, 0));
            b.redo(); assertEquals("7", b.raw(0, 1, 1));
        }
    }
    @Test void movedFormulaKeepsUnmovedDependenciesAndMovesInternalReferences() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "2"); b.set(0, 0, 1, "=A1+$D$1"); b.set(0, 0, 3, "10");
            CellMove.move(b, 0, 0, 0, 1, 2, 2, 1);
            assertEquals("=B3+$D$1", b.raw(0, 2, 2)); assertEquals("12", b.display(0, 2, 2));
        }
    }
    @Test void overlappingMovePreservesAllValuesAndRetargetsWholeRange() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "1"); b.set(0, 0, 1, "2"); b.set(0, 0, 2, "3");
            b.set(0, 2, 0, "=SUM(A1:C1)");
            CellMove.move(b, 0, 0, 0, 1, 3, 0, 1);
            assertEquals("", b.raw(0, 0, 0)); assertEquals("1", b.raw(0, 0, 1));
            assertEquals("2", b.raw(0, 0, 2)); assertEquals("3", b.raw(0, 0, 3));
            assertEquals("=SUM(B1:D1)", b.raw(0, 2, 0)); assertEquals("6", b.display(0, 2, 0));
        }
    }
    @Test void preservesLiteralBooleanErrorAndBlankTypes() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "'=text"); b.set(0, 0, 1, "TRUE"); b.cell(0, 0, 2, true).setCellErrorValue(FormulaError.NA.getCode());
            b.set(0, 1, 3, "overwritten");
            CellMove.move(b, 0, 0, 0, 1, 4, 1, 0);
            assertEquals(CellType.STRING, b.cell(0, 1, 0, false).getCellType()); assertEquals("=text", b.raw(0, 1, 0));
            assertEquals(CellType.BOOLEAN, b.cell(0, 1, 1, false).getCellType()); assertEquals("#N/A", b.raw(0, 1, 2));
            assertEquals("", b.raw(0, 1, 3));
        }
    }
    @Test void partialRangeRefusalLeavesOriginalAndDependenciesUnchanged() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "9"); b.set(0, 1, 0, "=A1"); b.set(0, 1, 1, "=SUM(A1:C1)");
            assertThrows(IllegalArgumentException.class, () -> b.transaction(x -> CellMove.move(x, 0, 0, 0, 1, 1, 3, 0)));
            assertEquals("9", b.raw(0, 0, 0)); assertEquals("=A1", b.raw(0, 1, 0)); assertEquals("", b.raw(0, 3, 0));
        }
    }
    @Test void boundaryAndMergedCellsAreRejectedWithoutDataLoss() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "9");
            assertThrows(IllegalArgumentException.class, () -> b.transaction(x -> CellMove.move(x, 0, 0, 0, 2, 1, Book.MAX_ROWS - 1, 0)));
            b.workbook().getSheetAt(0).addMergedRegion(CellRangeAddress.valueOf("A1:B1"));
            assertThrows(IllegalArgumentException.class, () -> b.transaction(x -> CellMove.move(x, 0, 0, 0, 1, 1, 3, 0)));
            assertEquals("9", b.raw(0, 0, 0));
        }
    }
    @Test void otherSheetLocalReferencesAreNotChangedAndDestinationReferencesStay() throws Exception {
        try (var b = new Book()) {
            b.workbook().createSheet("Other"); b.set(0, 0, 0, "5"); b.set(1, 0, 0, "=A2"); b.set(0, 4, 0, "=B2");
            CellMove.move(b, 0, 0, 0, 1, 1, 1, 1);
            assertEquals("=A2", b.raw(1, 0, 0)); assertEquals("=B2", b.raw(0, 4, 0)); assertEquals("5", b.display(0, 4, 0));
        }
    }
}
