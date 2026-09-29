package io.github.r2d2c2.gitsheet;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import java.nio.file.*;
import static org.junit.jupiter.api.Assertions.*;

class BookTest {
    @TempDir Path temp;
    @Test void evaluatesFunctionsAndInvalidatesDependencies() throws Exception {
        try (var book = new Book()) {
            book.set(0, 0, 0, "10"); book.set(0, 1, 0, "20"); book.set(0, 0, 1, "=SUM(A1:A2)");
            book.set(0, 1, 1, "=IF(B1>25,AVERAGE(A1:A2),0)");
            assertEquals("30", book.display(0, 0, 1)); assertEquals("15", book.display(0, 1, 1));
            book.set(0, 0, 0, "30"); assertEquals("50", book.display(0, 0, 1)); assertEquals("25", book.display(0, 1, 1));
        }
    }
    @Test void preservesTextBooleansLeadingZerosAndErrors() throws Exception {
        try (var book = new Book()) {
            book.set(0, 0, 0, "00123"); book.set(0, 0, 1, "TRUE"); book.set(0, 0, 2, "'=SUM(A1)"); book.set(0, 0, 3, "=1/0");
            assertEquals(CellType.STRING, book.cell(0, 0, 0, false).getCellType());
            assertEquals("TRUE", book.raw(0, 0, 1)); assertEquals("=SUM(A1)", book.raw(0, 0, 2)); assertEquals("#DIV/0!", book.display(0, 0, 3));
        }
    }
    @Test void supportsAtomicUndoRedoAndRollback() throws Exception {
        try (var book = new Book()) {
            book.transaction(b -> { b.set(0, 0, 0, "1"); b.set(0, 1, 0, "2"); });
            assertTrue(book.undo()); assertEquals("", book.raw(0, 0, 0)); assertEquals("", book.raw(0, 1, 0));
            assertTrue(book.redo()); assertEquals("2", book.raw(0, 1, 0));
            assertThrows(RuntimeException.class, () -> book.transaction(b -> { b.set(0, 0, 0, "99"); b.set(0, 2, 0, "=SUM("); }));
            assertEquals("1", book.raw(0, 0, 0)); assertEquals("", book.raw(0, 2, 0));
        }
    }
    @Test void editsAfterUndoDiscardRedoBranch() throws Exception {
        try (var book = new Book()) {
            book.transaction(b -> b.set(0, 0, 0, "1")); book.undo(); book.transaction(b -> b.set(0, 0, 0, "2"));
            assertFalse(book.redo()); assertEquals("2", book.raw(0, 0, 0));
        }
    }
    @Test void roundTripsDeterministicCanonicalDataAndStyles() throws Exception {
        var first = temp.resolve("first.gsheet"); var second = temp.resolve("second.gsheet");
        try (var book = new Book()) {
            book.set(0, 0, 0, "한글\n\t\"quoted\""); book.set(0, 1, 0, "1234.5");
            book.set(0, 2, 0, "=A2*2"); book.set(0, 3, 0, "TRUE"); book.set(0, 4, 0, "'=literal");
            book.style(0, 1, 0, "format", "#,##0.00"); book.style(0, 1, 0, "bold", "true");
            book.style(0, 1, 0, "fill", Short.toString(IndexedColors.LIGHT_YELLOW.getIndex()));
            var sheet = book.workbook().getSheetAt(0); sheet.setColumnWidth(0, 5000); sheet.createFreezePane(1, 1);
            sheet.addMergedRegion(new CellRangeAddress(8, 8, 0, 1));
            BookFiles.write(book, first); BookFiles.write(book, second); assertEquals(Files.readString(first), Files.readString(second));
        }
        try (var loaded = BookFiles.read(first)) {
            assertEquals("한글\n\t\"quoted\"", loaded.raw(0, 0, 0)); assertEquals("2469", loaded.display(0, 2, 0));
            assertEquals("=literal", loaded.raw(0, 4, 0)); assertEquals(CellType.STRING, loaded.cell(0, 4, 0, false).getCellType());
            assertEquals("1,234.50", loaded.display(0, 1, 0)); assertTrue(loaded.workbook().getFontAt(loaded.cell(0, 1, 0, false).getCellStyle().getFontIndex()).getBold());
            assertEquals(5000, loaded.workbook().getSheetAt(0).getColumnWidth(0)); assertEquals(1, loaded.workbook().getSheetAt(0).getNumMergedRegions());
            BookFiles.write(loaded, second); assertEquals(Files.readString(first), Files.readString(second));
        }
    }
    @Test void forwardCrossSheetReferencesSurviveCanonicalReload() throws Exception {
        var path = temp.resolve("cross.gsheet");
        try (var book = new Book()) {
            book.workbook().createSheet("다음 시트"); book.set(1, 0, 0, "42"); book.set(0, 0, 0, "='다음 시트'!A1+1"); BookFiles.write(book, path);
        }
        try (var loaded = BookFiles.read(path)) { assertEquals("43", loaded.display(0, 0, 0)); }
    }
    @Test void excelRoundTripPreservesFormulaAndFormat() throws Exception {
        var path = temp.resolve("book.xlsx");
        try (var book = new Book()) {
            book.set(0, 0, 0, "0.25"); book.style(0, 0, 0, "format", "0.00%"); book.set(0, 0, 1, "=A1*4"); BookFiles.exportExcel(book, path);
        }
        try (var loaded = BookFiles.importExcel(path)) { assertEquals("25.00%", loaded.display(0, 0, 0)); assertEquals("=A1*4", loaded.raw(0, 0, 1)); assertEquals("1", loaded.display(0, 0, 1)); }
    }
    @Test void csvHandlesQuotedCommasNewlinesAndUntrustedFormulas() throws Exception {
        var csv = temp.resolve("input.csv"); Files.writeString(csv, "\"a,b\",\"two\nlines\",=1+1\r\n한글,0012,TRUE\r\n");
        try (var book = BookFiles.importCsv(csv)) {
            assertEquals("a,b", book.raw(0, 0, 0)); assertEquals("two\nlines", book.raw(0, 0, 1));
            assertEquals(CellType.STRING, book.cell(0, 0, 2, false).getCellType());
            assertEquals("0012", book.raw(0, 1, 1)); var out = temp.resolve("out.csv"); BookFiles.exportCsv(book, 0, out);
            try (var second = BookFiles.importCsv(out)) { assertEquals("two\nlines", second.raw(0, 0, 1)); }
        }
    }
    @Test void rejectsUnknownOrCorruptFormat() throws Exception {
        var path = temp.resolve("bad.gsheet"); Files.writeString(path, "{\"type\":\"git-sheet\",\"version\":99}\n");
        assertThrows(java.io.IOException.class, () -> BookFiles.read(path)); Files.writeString(path, "not json");
        assertThrows(java.io.IOException.class, () -> BookFiles.read(path));
    }
    @Test void enforcesExcelCoordinatesWithoutAllocatingEmptyRows() throws Exception {
        try (var book = new Book()) {
            assertNull(book.cell(0, Book.MAX_ROWS - 1, Book.MAX_COLUMNS - 1, false));
            assertEquals(0, book.workbook().getSheetAt(0).getPhysicalNumberOfRows());
            assertThrows(IllegalArgumentException.class, () -> book.set(0, Book.MAX_ROWS, 0, "x"));
            assertEquals("XFD1048576", Book.address(Book.MAX_ROWS - 1, Book.MAX_COLUMNS - 1));
        }
    }
    @Test void cachesStylesAcrossRange() throws Exception {
        try (var book = new Book()) {
            for (int r = 0; r < 100; r++) book.style(0, r, 0, "bold", "true");
            assertTrue(book.workbook().getNumCellStyles() < 5);
        }
    }
    @Test void singleCellChangeProducesOneChangedCanonicalLine() throws Exception {
        try (var book = new Book()) {
            for (int r = 0; r < 10; r++) book.set(0, r, 0, Integer.toString(r));
            var path = temp.resolve("book.gsheet"); BookFiles.write(book, path); var before = Files.readAllLines(path);
            book.set(0, 5, 0, "99"); BookFiles.write(book, path); var after = Files.readAllLines(path);
            long changed = java.util.stream.IntStream.range(0, before.size()).filter(i -> !before.get(i).equals(after.get(i))).count();
            assertEquals(1, changed);
        }
    }
}
