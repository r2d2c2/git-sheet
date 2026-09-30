package io.github.r2d2c2.gitsheet;

import java.nio.file.Path;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.apache.poi.ss.util.CellRangeAddress;
import static org.junit.jupiter.api.Assertions.*;

class MergeEditsTest {
    @Test void refusesExcelTableIntersections() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "Header1"); b.set(0, 0, 1, "Header2");
            b.workbook().getSheetAt(0).createTable(new org.apache.poi.ss.util.AreaReference("A1:B3", org.apache.poi.ss.SpreadsheetVersion.EXCEL2007));
            assertThrows(IllegalArgumentException.class, () -> MergeEdits.merge(b, 0, CellRangeAddress.valueOf("A2:B3")));
            assertEquals(0, b.workbook().getSheetAt(0).getNumMergedRegions());
        }
    }
    @Test void mergeUnmergeUndoAndFileRoundTrip(@TempDir Path dir) throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "Title"); b.style(0, 1, 1, "bold", "true");
            b.transaction(x -> MergeEdits.merge(x, 0, CellRangeAddress.valueOf("A1:B2")));
            assertEquals("A1:B2", MergeEdits.containing(b.workbook().getSheetAt(0), 1, 1).formatAsString());
            for (String ext : new String[]{"gsheet", "xlsx"}) {
                var path = dir.resolve("merged." + ext);
                if (ext.equals("xlsx")) BookFiles.exportExcel(b, path); else BookFiles.write(b, path);
                try (var loaded = ext.equals("xlsx") ? BookFiles.importExcel(path) : BookFiles.read(path)) {
                    assertEquals("A1:B2", loaded.workbook().getSheetAt(0).getMergedRegion(0).formatAsString());
                    assertEquals("Title", loaded.raw(0, 0, 0));
                }
            }
            b.transaction(x -> MergeEdits.unmerge(x, 0, CellRangeAddress.valueOf("B2")));
            assertEquals(0, b.workbook().getSheetAt(0).getNumMergedRegions());
            assertTrue(b.workbook().getFontAt(b.cell(0, 1, 1, false).getCellStyle().getFontIndex()).getBold());
            b.undo(); assertEquals(1, b.workbook().getSheetAt(0).getNumMergedRegions());
            b.undo(); assertEquals(0, b.workbook().getSheetAt(0).getNumMergedRegions());
            b.redo(); assertEquals(1, b.workbook().getSheetAt(0).getNumMergedRegions());
        }
    }
    @Test void refusesCoveredValuesIncludingEmptyStringsAndFormulas() throws Exception {
        try (var b = new Book()) {
            for (String value : new String[]{"text", "0", "FALSE", "=1+1", "'"}) {
                b.set(0, 0, 1, value);
                assertThrows(IllegalArgumentException.class, () -> MergeEdits.merge(b, 0, CellRangeAddress.valueOf("A1:B2")));
                assertEquals(0, b.workbook().getSheetAt(0).getNumMergedRegions());
            }
        }
    }
    @Test void overlapProtectionArraysAndSingleCellsAreRefused() throws Exception {
        try (var b = new Book()) {
            assertThrows(IllegalArgumentException.class, () -> MergeEdits.merge(b, 0, CellRangeAddress.valueOf("A1")));
            MergeEdits.merge(b, 0, CellRangeAddress.valueOf("A1:B2"));
            MergeEdits.merge(b, 0, CellRangeAddress.valueOf("A1:B2"));
            assertEquals(1, b.workbook().getSheetAt(0).getNumMergedRegions());
            assertThrows(IllegalArgumentException.class, () -> MergeEdits.merge(b, 0, CellRangeAddress.valueOf("B2:C3")));
            b.workbook().getSheetAt(0).setArrayFormula("ROW(D1:D2)", CellRangeAddress.valueOf("D1:D2"));
            assertThrows(IllegalArgumentException.class, () -> MergeEdits.merge(b, 0, CellRangeAddress.valueOf("D1:E2")));
            b.workbook().getSheetAt(0).protectSheet("test");
            assertThrows(IllegalArgumentException.class, () -> MergeEdits.merge(b, 0, CellRangeAddress.valueOf("F1:G2")));
            assertThrows(IllegalArgumentException.class, () -> MergeEdits.unmerge(b, 0, CellRangeAddress.valueOf("A1")));
        }
    }
    @Test void unmergeImportedHiddenValuesNeverDeletesThem() throws Exception {
        try (var b = new Book()) {
            b.set(0, 0, 0, "anchor"); b.set(0, 1, 1, "hidden");
            b.workbook().getSheetAt(0).addMergedRegion(CellRangeAddress.valueOf("A1:B2"));
            MergeEdits.unmerge(b, 0, CellRangeAddress.valueOf("B2"));
            assertEquals("anchor", b.raw(0, 0, 0)); assertEquals("hidden", b.raw(0, 1, 1));
        }
    }
}
