import java.io.BufferedReader;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.LinkedHashMap;
import java.util.Map;
import java.math.BigDecimal;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public final class vba_sample_3 {
    private vba_sample_3() { }

    public static void ImportDailySales(
        Path inputWorkbook, Path tsvPath, LocalDate reportDate, Path outputWorkbook
    ) throws IOException {
        if (!Files.isRegularFile(tsvPath)) {
            throw new IOException(
                "ImportDailySales error -2147219503: Input file does not exist: " + tsvPath
            );
        }
        Map<String, BigDecimal> salesByRegion = new LinkedHashMap<>();
        try (BufferedReader stream = Files.newBufferedReader(tsvPath, StandardCharsets.UTF_8)) {
            stream.readLine();
            String lineText;
            while ((lineText = stream.readLine()) != null) {
                String[] fields = lineText.split("\\t", -1);
                if (fields.length >= 3) {
                    String regionCode = VbaExcelSupport.normalizeCode(fields[1]);
                    if (!regionCode.isEmpty() && isNumeric(fields[2])) {
                        BigDecimal saleAmount = VbaExcelSupport.currency(new BigDecimal(fields[2].trim()));
                        BigDecimal total = salesByRegion.getOrDefault(regionCode, BigDecimal.ZERO)
                            .add(saleAmount);
                        salesByRegion.put(regionCode, VbaExcelSupport.currency(total));
                    }
                }
            }
        }

        try (Workbook workbook = VbaExcelSupport.open(inputWorkbook)) {
            Sheet sheet = workbook.getSheet("Daily Sales");
            if (sheet == null) {
                throw new IllegalArgumentException("Required worksheet not found: Daily Sales");
            }
            int lastRow = VbaExcelSupport.lastRowInColumn(sheet, 0);
            for (int row = 1; row <= lastRow; row++) {
                Cell dateCell = VbaExcelSupport.cell(sheet, row, 0);
                if (VbaExcelSupport.isDate(dateCell) &&
                        VbaExcelSupport.dateTime(dateCell).toLocalDate().equals(reportDate)) {
                    String regionCode = VbaExcelSupport.normalizeCode(
                        VbaExcelSupport.text(VbaExcelSupport.cell(sheet, row, 1))
                    );
                    BigDecimal amount = salesByRegion.getOrDefault(regionCode, BigDecimal.ZERO);
                    VbaExcelSupport.set(VbaExcelSupport.writableCell(sheet, row, 2), amount);
                }
            }
            VbaExcelSupport.writeAtomically(workbook, outputWorkbook);
        }
    }

    private static boolean isNumeric(String value) {
        try {
            new BigDecimal(value.trim());
            return true;
        } catch (NumberFormatException exception) {
            return false;
        }
    }

    public static void main(String[] args) throws IOException {
        if (args.length < 3 || args.length > 4) {
            System.err.println(
                "Usage: java vba_sample_3 <workbook.xlsx> <sales.tsv> <report-date> [output.xlsx]"
            );
            System.exit(2);
        }
        Path inputWorkbook = Path.of(args[0]);
        Path outputWorkbook = args.length == 4
            ? Path.of(args[3]) : VbaExcelSupport.defaultOutput(inputWorkbook);
        ImportDailySales(
            inputWorkbook, Path.of(args[1]), VbaExcelSupport.parseDate(args[2]), outputWorkbook
        );
    }
}
