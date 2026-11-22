import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.ZoneId;
import java.util.Locale;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.DateUtil;

final class VbaExcelSupport {
    static final DataFormatter FORMATTER = new DataFormatter(Locale.getDefault());

    private VbaExcelSupport() { }

    static Workbook open(Path path) throws IOException {
        return WorkbookFactory.create(path.toFile());
    }

    static Path defaultOutput(Path input) {
        String fileName = input.getFileName().toString();
        int dot = fileName.lastIndexOf('.');
        String stem = dot > 0 ? fileName.substring(0, dot) : fileName;
        return input.resolveSibling(stem + "_java.xlsx");
    }

    static void writeAtomically(Workbook workbook, Path output) throws IOException {
        Path absolute = output.toAbsolutePath().normalize();
        Path parent = absolute.getParent();
        if (parent != null) Files.createDirectories(parent);
        Path temporary = Files.createTempFile(parent, ".vba-java-", ".xlsx");
        boolean moved = false;
        try {
            try (OutputStream stream = Files.newOutputStream(temporary)) {
                workbook.write(stream);
            }
            try {
                Files.move(temporary, absolute, StandardCopyOption.REPLACE_EXISTING,
                        StandardCopyOption.ATOMIC_MOVE);
            } catch (AtomicMoveNotSupportedException exception) {
                Files.move(temporary, absolute, StandardCopyOption.REPLACE_EXISTING);
            }
            moved = true;
        } finally {
            if (!moved) Files.deleteIfExists(temporary);
        }
    }

    static Cell cell(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        return row == null ? null : row.getCell(columnIndex);
    }

    static Cell writableCell(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) row = sheet.createRow(rowIndex);
        Cell cell = row.getCell(columnIndex);
        return cell == null ? row.createCell(columnIndex) : cell;
    }

    static String text(Cell cell) {
        if (cell == null) return "";
        if (cell.getCellType() != CellType.FORMULA) return FORMATTER.formatCellValue(cell);
        return switch (cell.getCachedFormulaResultType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> FORMATTER.formatRawCellContents(
                cell.getNumericCellValue(),
                cell.getCellStyle().getDataFormat(),
                cell.getCellStyle().getDataFormatString()
            );
            case BOOLEAN -> Boolean.toString(cell.getBooleanCellValue());
            case BLANK, _NONE, ERROR, FORMULA -> "";
        };
    }

    static Object value(Cell cell) {
        if (cell == null) return null;
        CellType type = cell.getCellType() == CellType.FORMULA
                ? cell.getCachedFormulaResultType() : cell.getCellType();
        return switch (type) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> DateUtil.isCellDateFormatted(cell)
                    ? cell.getLocalDateTimeCellValue() : cell.getNumericCellValue();
            case BOOLEAN -> cell.getBooleanCellValue();
            case BLANK, _NONE, ERROR, FORMULA -> null;
        };
    }

    static void copyValue(Cell from, Cell to) {
        Object value = value(from);
        if (value == null) {
            to.setBlank();
        } else if (value instanceof String text) {
            to.setCellValue(text);
        } else if (value instanceof Boolean bool) {
            to.setCellValue(bool);
        } else if (value instanceof LocalDateTime dateTime) {
            to.setCellValue(dateTime);
        } else if (value instanceof Number number) {
            to.setCellValue(number.doubleValue());
        } else {
            to.setCellValue(value.toString());
        }
    }

    static void set(Cell cell, Object value) {
        if (value == null) {
            cell.setBlank();
        } else if (value instanceof String text) {
            cell.setCellValue(text);
        } else if (value instanceof Boolean bool) {
            cell.setCellValue(bool);
        } else if (value instanceof BigDecimal decimal) {
            cell.setCellValue(decimal.doubleValue());
        } else if (value instanceof Number number) {
            cell.setCellValue(number.doubleValue());
        } else if (value instanceof LocalDateTime dateTime) {
            cell.setCellValue(dateTime);
        } else if (value instanceof LocalDate date) {
            cell.setCellValue(date.atStartOfDay());
        } else {
            cell.setCellValue(value.toString());
        }
    }

    static boolean isNumeric(Cell cell) {
        if (cell == null) return false;
        if (value(cell) instanceof Number) return true;
        try {
            new BigDecimal(text(cell).trim());
            return true;
        } catch (NumberFormatException exception) {
            return false;
        }
    }

    static BigDecimal decimal(Cell cell) {
        if (cell == null) throw new IllegalArgumentException("Expected a numeric cell");
        Object value = value(cell);
        if (value instanceof Number number) return BigDecimal.valueOf(number.doubleValue());
        return new BigDecimal(text(cell).trim());
    }

    static BigDecimal currency(BigDecimal amount) {
        BigDecimal rounded = amount.setScale(4, RoundingMode.HALF_EVEN);
        if (rounded.compareTo(new BigDecimal("922337203685477.5807")) > 0 ||
                rounded.compareTo(new BigDecimal("-922337203685477.5808")) < 0) {
            throw new ArithmeticException("Currency value is outside the VBA Currency range");
        }
        return rounded;
    }

    static LocalDateTime dateTime(Cell cell) {
        if (cell == null) return null;
        CellType type = cell.getCellType() == CellType.FORMULA
                ? cell.getCachedFormulaResultType() : cell.getCellType();
        if (type == CellType.NUMERIC) {
            return DateUtil.getJavaDate(cell.getNumericCellValue()).toInstant()
                    .atZone(ZoneId.systemDefault()).toLocalDateTime();
        }
        return parseDateTime(text(cell));
    }

    private static LocalDateTime parseDateTime(String text) {
        String value = text.trim();
        for (DateTimeFormatter formatter : new DateTimeFormatter[] {
                DateTimeFormatter.ISO_LOCAL_DATE_TIME,
                DateTimeFormatter.ofPattern("M/d/uuuu H:mm:ss"),
                DateTimeFormatter.ofPattern("M/d/uuuu H:mm")}) {
            try {
                return LocalDateTime.parse(value, formatter);
            } catch (DateTimeParseException ignored) { }
        }
        return parseDate(value).atStartOfDay();
    }

    static LocalDate parseDate(String text) {
        String value = text.trim();
        if (value.isEmpty()) throw new DateTimeParseException("Empty date", value, 0);
        try {
            return LocalDateTime.parse(value, DateTimeFormatter.ISO_LOCAL_DATE_TIME).toLocalDate();
        } catch (DateTimeParseException ignored) { }
        for (DateTimeFormatter formatter : new DateTimeFormatter[] {
                DateTimeFormatter.ISO_LOCAL_DATE,
                DateTimeFormatter.ofPattern("M/d/uuuu"),
                DateTimeFormatter.ofPattern("M/d/uu"),
                DateTimeFormatter.ofPattern("d-MMM-uuuu", Locale.ENGLISH)}) {
            try {
                return LocalDate.parse(value, formatter);
            } catch (DateTimeParseException ignored) { }
        }
        throw new DateTimeParseException("Unsupported VBA date value", value, 0);
    }

    static boolean isDate(Cell cell) {
        try {
            dateTime(cell);
            return cell != null;
        } catch (RuntimeException exception) {
            return false;
        }
    }

    static int lastRowInColumn(Sheet sheet, int columnIndex) {
        for (int row = sheet.getLastRowNum(); row >= 0; row--) {
            Cell cell = cell(sheet, row, columnIndex);
            if (cell == null) continue;
            CellType type = cell.getCellType();
            if (type == CellType.FORMULA || type == CellType.NUMERIC ||
                    type == CellType.BOOLEAN || type == CellType.ERROR) return row;
            if (type == CellType.STRING && !cell.getStringCellValue().isEmpty()) return row;
        }
        return 0;
    }

    static void fill(Cell cell, IndexedColors color) {
        Workbook workbook = cell.getSheet().getWorkbook();
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(cell.getCellStyle());
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(style);
    }

    static void autoSize(Sheet sheet, int columns) {
        for (int column = 0; column < columns; column++) {
            sheet.autoSizeColumn(column);
        }
    }

    static String normalizeCode(String text) {
        return text.trim().toUpperCase(Locale.getDefault());
    }

    static String excelColumn(int zeroBasedColumn) {
        int number = zeroBasedColumn + 1;
        StringBuilder result = new StringBuilder();
        while (number > 0) {
            int remainder = (number - 1) % 26;
            result.insert(0, (char) ('A' + remainder));
            number = (number - 1) / 26;
        }
        return result.toString();
    }
}
