package io.github.r2d2c2.gitsheet;

import java.nio.file.*;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import static org.junit.jupiter.api.Assertions.*;

class RowLayoutTest {
    @Test void manualAndAutomaticHeightParticipateInUndo() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "first\nsecond");
            b.transaction(x -> { x.setRowHeight(0, 0, 45.0); x.style(0, 0, 0, "wrap", "true"); });
            assertEquals(45, b.workbook().getSheetAt(0).getRow(0).getHeightInPoints());
            assertTrue(b.cell(0, 0, 0, false).getCellStyle().getWrapText());
            b.transaction(x -> x.setRowHeight(0, 0, null)); assertFalse(b.workbook().getSheetAt(0).getRow(0).getCTRow().isSetHt());
            b.undo(); assertEquals(45, b.workbook().getSheetAt(0).getRow(0).getHeightInPoints());
            b.undo(); assertFalse(b.cell(0, 0, 0, false).getCellStyle().getWrapText());
            assertEquals("first\nsecond", b.raw(0, 0, 0)); b.redo(); assertEquals(45, b.workbook().getSheetAt(0).getRow(0).getHeightInPoints());
        }
    }
    @Test void layoutRoundTripsIncludingExplicitDefaultHeightAndHiddenAutoRow(@TempDir Path dir) throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "line 1\nline 2"); b.style(0, 0, 0, "wrap", "true"); b.setRowHeight(0, 0, 60.0);
            b.setRowHeight(0, 1, (double) b.workbook().getSheetAt(0).getDefaultRowHeightInPoints());
            b.workbook().getSheetAt(0).createRow(2).setZeroHeight(true);
            var text = dir.resolve("layout.gsheet"); BookFiles.write(b, text);
            try (var loaded = BookFiles.read(text)) { check(loaded); }
            var excel = dir.resolve("layout.xlsx"); BookFiles.exportExcel(b, excel);
            try (var loaded = BookFiles.importExcel(excel)) { check(loaded); }
        }
    }
    private static void check(Book b) {
        assertEquals(60, b.workbook().getSheetAt(0).getRow(0).getHeightInPoints());
        assertTrue(b.cell(0, 0, 0, false).getCellStyle().getWrapText()); assertEquals("line 1\nline 2", b.raw(0, 0, 0));
        assertTrue(b.workbook().getSheetAt(0).getRow(1).getCTRow().isSetHt());
        assertTrue(b.workbook().getSheetAt(0).getRow(2).getZeroHeight()); assertFalse(b.workbook().getSheetAt(0).getRow(2).getCTRow().isSetHt());
    }
    @Test void invalidHeightDoesNotCreateRowsAndAutoResetPreservesSparseSheets() throws Exception {
        try (var b = new Book()) {
            for (double value : new double[] { 0, -1, 410, Double.NaN, Double.POSITIVE_INFINITY })
                assertThrows(IllegalArgumentException.class, () -> b.setRowHeight(0, 500, value));
            b.setRowHeight(0, 500, null); assertNull(b.workbook().getSheetAt(0).getRow(500));
        }
    }
}
