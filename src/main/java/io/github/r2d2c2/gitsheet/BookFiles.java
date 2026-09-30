package io.github.r2d2c2.gitsheet;

import com.google.gson.*;
import org.apache.commons.csv.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;

/** Stable, UTF-8 JSON Lines: one record per sheet, row layout and nonempty/styled cell. */
public final class BookFiles {
    private static final Gson JSON = new GsonBuilder().disableHtmlEscaping().create();
    private BookFiles() {}
    public record Style(String font, short size, boolean bold, boolean italic, String color,
                        String fill, String format, String align, boolean wrap, String borderTop,
                        String borderBottom, String borderLeft, String borderRight) {}

    public static void write(Book book, Path path) throws IOException {
        atomic(path, out -> {
            var writer = new BufferedWriter(new OutputStreamWriter(out, StandardCharsets.UTF_8));
            var header = record("git-sheet"); header.addProperty("version", 1); line(writer, header);
            for (var sheet : book.workbook()) {
                var record = record("sheet"); record.addProperty("name", sheet.getSheetName());
                record.addProperty("defaultColumnWidth", sheet.getDefaultColumnWidth());
                record.addProperty("defaultRowHeight", sheet.getDefaultRowHeight());
                var widths = new TreeMap<Integer, Integer>();
                int lastColumn = 0;
                for (var row : sheet) lastColumn = Math.max(lastColumn, row.getLastCellNum());
                for (var group : ((XSSFSheet) sheet).getCTWorksheet().getColsList())
                    for (var column : group.getColList()) lastColumn = Math.max(lastColumn, (int) column.getMax());
                var hiddenColumns = new ArrayList<Integer>();
                for (int c = 0; c < lastColumn; c++) {
                    if (sheet.getColumnWidth(c) != sheet.getDefaultColumnWidth() * 256) widths.put(c, sheet.getColumnWidth(c));
                    if (sheet.isColumnHidden(c)) hiddenColumns.add(c);
                }
                record.add("widths", JSON.toJsonTree(widths));
                if (!hiddenColumns.isEmpty()) record.add("hiddenColumns", JSON.toJsonTree(hiddenColumns));
                var merges = sheet.getMergedRegions().stream().map(CellRangeAddress::formatAsString).toList();
                record.add("merges", JSON.toJsonTree(merges));
                var pane = sheet.getPaneInformation();
                if (pane != null && pane.isFreezePane()) {
                    record.addProperty("freezeRows", pane.getHorizontalSplitPosition());
                    record.addProperty("freezeColumns", pane.getVerticalSplitPosition());
                }
                line(writer, record);
                for (var row : sheet) {
                    if (((XSSFRow) row).getCTRow().isSetHt() || row.getZeroHeight()) {
                        var r = record("row"); r.addProperty("row", row.getRowNum()); r.addProperty("height", row.getHeight());
                        r.addProperty("customHeight", ((XSSFRow) row).getCTRow().isSetHt());
                        r.addProperty("hidden", row.getZeroHeight()); line(writer, r);
                    }
                    for (var cell : row) {
                        if (cell.getCellType() == CellType.BLANK && cell.getCellStyle().getIndex() == 0) continue;
                        var c = record("cell"); c.addProperty("address", cell.getAddress().formatAsString());
                        c.addProperty("kind", cell.getCellType().name());
                        c.addProperty("value", book.raw(book.workbook().getSheetIndex(sheet), row.getRowNum(), cell.getColumnIndex()));
                        c.add("style", JSON.toJsonTree(style(book.workbook(), cell.getCellStyle())));
                        line(writer, c);
                    }
                }
            }
            writer.flush();
        });
    }
    public static Book read(Path path) throws IOException {
        var workbook = new XSSFWorkbook();
        try (var reader = Files.newBufferedReader(path, StandardCharsets.UTF_8)) {
            var header = JsonParser.parseString(Objects.requireNonNull(reader.readLine(), "빈 문서")).getAsJsonObject();
            if (!header.get("type").getAsString().equals("git-sheet") || header.get("version").getAsInt() != 1)
                throw new IOException("지원하지 않는 Git Sheet 파일 형식입니다.");
            XSSFSheet sheet = null;
            var styles = new HashMap<Style, CellStyle>();
            String line;
            while ((line = reader.readLine()) != null) {
                if (line.isBlank()) continue;
                var item = JsonParser.parseString(line).getAsJsonObject();
                switch (item.get("type").getAsString()) {
                    case "sheet" -> {
                        sheet = workbook.createSheet(item.get("name").getAsString());
                        sheet.setDefaultColumnWidth(item.get("defaultColumnWidth").getAsInt());
                        sheet.setDefaultRowHeight(item.get("defaultRowHeight").getAsShort());
                        for (var entry : item.getAsJsonObject("widths").entrySet()) sheet.setColumnWidth(Integer.parseInt(entry.getKey()), entry.getValue().getAsInt());
                        if (item.has("hiddenColumns")) for (var column : item.getAsJsonArray("hiddenColumns")) sheet.setColumnHidden(column.getAsInt(), true);
                        for (var merge : item.getAsJsonArray("merges")) sheet.addMergedRegion(CellRangeAddress.valueOf(merge.getAsString()));
                        if (item.has("freezeRows")) sheet.createFreezePane(item.get("freezeColumns").getAsInt(), item.get("freezeRows").getAsInt());
                    }
                    case "row" -> {
                        Objects.requireNonNull(sheet, "시트가 없습니다.");
                        var row = sheet.createRow(item.get("row").getAsInt());
                        if (!item.has("customHeight") || item.get("customHeight").getAsBoolean()) row.setHeight(item.get("height").getAsShort());
                        row.setZeroHeight(item.get("hidden").getAsBoolean());
                    }
                    case "cell" -> {
                        Objects.requireNonNull(sheet, "시트가 없습니다.");
                        var address = new org.apache.poi.ss.util.CellAddress(item.get("address").getAsString());
                        var row = sheet.getRow(address.getRow()); if (row == null) row = sheet.createRow(address.getRow());
                        var cell = row.createCell(address.getColumn()); var value = item.get("value").getAsString();
                        switch (CellType.valueOf(item.get("kind").getAsString())) {
                            case STRING -> cell.setCellValue(value);
                            case NUMERIC -> cell.setCellValue(Double.parseDouble(value));
                            case BOOLEAN -> cell.setCellValue(Boolean.parseBoolean(value));
                            case FORMULA -> cell.setCellFormula(value.substring(1));
                            case ERROR -> cell.setCellErrorValue(FormulaError.forString(value).getCode());
                            default -> cell.setBlank();
                        }
                        var style = JSON.fromJson(item.get("style"), Style.class);
                        cell.setCellStyle(styles.computeIfAbsent(style, key -> makeStyle(workbook, key)));
                    }
                    default -> throw new IOException("알 수 없는 레코드입니다.");
                }
            }
            if (workbook.getNumberOfSheets() == 0) throw new IOException("시트가 없습니다.");
            // Formula validation/evaluation is deferred until every referenced sheet has been loaded.
            return new Book(workbook);
        } catch (Exception e) {
            workbook.close();
            if (e instanceof IOException io) throw io;
            throw new IOException("문서를 읽을 수 없습니다: " + e.getMessage(), e);
        }
    }
    private static Style style(XSSFWorkbook book, CellStyle cs) {
        var s = (XSSFCellStyle) cs; var font = book.getFontAt(s.getFontIndex());
        return new Style(font.getFontName(), font.getFontHeight(), font.getBold(), font.getItalic(),
                color(font.getXSSFColor(), font.getColor()),
                s.getFillPattern() == FillPatternType.SOLID_FOREGROUND ? color(s.getFillForegroundXSSFColor(), s.getFillForegroundColor()) : "none",
                s.getDataFormatString(), s.getAlignment().name(), s.getWrapText(), s.getBorderTop().name(),
                s.getBorderBottom().name(), s.getBorderLeft().name(), s.getBorderRight().name());
    }
    private static String color(XSSFColor color, short index) {
        return color != null && color.getARGBHex() != null ? "#" + color.getARGBHex() : "index:" + index;
    }
    private static XSSFColor parseColor(String value) {
        var color = new XSSFColor(new DefaultIndexedColorMap());
        if (value.startsWith("#")) color.setARGBHex(value.substring(1)); else color.setIndexed(Integer.parseInt(value.substring(6)));
        return color;
    }
    private static CellStyle makeStyle(XSSFWorkbook book, Style value) {
        var s = book.createCellStyle(); var font = book.createFont();
        font.setFontName(value.font()); font.setFontHeight(value.size()); font.setBold(value.bold()); font.setItalic(value.italic());
        font.setColor(parseColor(value.color())); s.setFont(font);
        if (!value.fill().equals("none")) { s.setFillForegroundColor(parseColor(value.fill())); s.setFillPattern(FillPatternType.SOLID_FOREGROUND); }
        s.setDataFormat(book.createDataFormat().getFormat(value.format())); s.setAlignment(HorizontalAlignment.valueOf(value.align()));
        s.setWrapText(value.wrap()); s.setBorderTop(BorderStyle.valueOf(value.borderTop())); s.setBorderBottom(BorderStyle.valueOf(value.borderBottom()));
        s.setBorderLeft(BorderStyle.valueOf(value.borderLeft())); s.setBorderRight(BorderStyle.valueOf(value.borderRight())); return s;
    }
    private static JsonObject record(String type) { var o = new JsonObject(); o.addProperty("type", type); return o; }
    private static void line(Writer writer, Object value) throws IOException { writer.write(JSON.toJson(value)); writer.write('\n'); }
    public static Book importExcel(Path path) throws IOException {
        try (var stream = Files.newInputStream(path)) { return new Book(new XSSFWorkbook(stream)); }
    }
    public static void exportExcel(Book book, Path path) throws IOException {
        book.workbook().setForceFormulaRecalculation(true);
        atomic(path, out -> book.workbook().write(out));
    }
    public static Book importCsv(Path path) throws IOException {
        var book = new Book();
        try (var reader = Files.newBufferedReader(path, StandardCharsets.UTF_8);
             var parser = CSVFormat.DEFAULT.parse(reader)) {
            int r = 0;
            for (var row : parser) {
                for (int c = 0; c < row.size(); c++) {
                    // CSV is data: formulas from external CSV are never executed on import.
                    book.cell(0, r, c, true).setCellValue(row.get(c));
                }
                r++;
            }
            return book;
        } catch (Exception e) { book.close(); throw e; }
    }
    public static void exportCsv(Book book, int sheetIndex, Path path) throws IOException {
        atomic(path, out -> {
            var writer = new OutputStreamWriter(out, StandardCharsets.UTF_8);
            var csv = new CSVPrinter(writer, CSVFormat.DEFAULT);
            var sheet = book.workbook().getSheetAt(sheetIndex);
            int columns = 0;
            for (var row : sheet) columns = Math.max(columns, row.getLastCellNum());
            for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                var values = new ArrayList<String>();
                for (int c = 0; c < columns; c++) values.add(book.display(sheetIndex, r, c));
                csv.printRecord(values);
            }
            csv.flush();
        });
    }
    @FunctionalInterface private interface WriteAction { void write(OutputStream out) throws IOException; }
    private static void atomic(Path path, WriteAction write) throws IOException {
        path = path.toAbsolutePath(); Files.createDirectories(path.getParent());
        var temporary = Files.createTempFile(path.getParent(), ".git-sheet-", ".tmp");
        try {
            try (var out = Files.newOutputStream(temporary)) { write.write(out); }
            try { Files.move(temporary, path, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE); }
            catch (AtomicMoveNotSupportedException e) { Files.move(temporary, path, StandardCopyOption.REPLACE_EXISTING); }
        } finally { Files.deleteIfExists(temporary); }
    }
}
