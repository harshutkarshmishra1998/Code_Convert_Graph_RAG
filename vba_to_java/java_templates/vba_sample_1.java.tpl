import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDateTime;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public final class vba_sample_1 {
    private vba_sample_1() { }

    public static void RefreshMonthlySummary(Path inputWorkbook, Path outputWorkbook) throws IOException {
        try (Workbook workbook = VbaExcelSupport.open(inputWorkbook)) {
            Sheet sourceSheet = requireSheet(workbook, "Transactions");
            Sheet reportSheet = requireSheet(workbook, "Monthly Summary");
            int lastRow = VbaExcelSupport.lastRowInColumn(sourceSheet, 0);
            int reportRow = 1;

            for (int rowNumber = 1; rowNumber <= lastRow; rowNumber++) {
                Cell account = VbaExcelSupport.cell(sourceSheet, rowNumber, 0);
                if (!VbaExcelSupport.text(account).trim().isEmpty()) {
                    Cell dateCell = VbaExcelSupport.cell(sourceSheet, rowNumber, 1);
                    if (VbaExcelSupport.isDate(dateCell)) {
                        VbaExcelSupport.copyValue(
                            account, VbaExcelSupport.writableCell(reportSheet, reportRow, 0)
                        );
                        LocalDateTime date = VbaExcelSupport.dateTime(dateCell);
                        VbaExcelSupport.set(
                            VbaExcelSupport.writableCell(reportSheet, reportRow, 1), date
                        );
                        VbaExcelSupport.copyValue(
                            VbaExcelSupport.cell(sourceSheet, rowNumber, 2),
                            VbaExcelSupport.writableCell(reportSheet, reportRow, 2)
                        );
                        reportRow++;
                    } else {
                        VbaExcelSupport.fill(
                            VbaExcelSupport.writableCell(sourceSheet, rowNumber, 1),
                            IndexedColors.YELLOW
                        );
                    }
                }
            }

            VbaExcelSupport.autoSize(reportSheet, 3);
            VbaExcelSupport.writeAtomically(workbook, outputWorkbook);
            System.out.printf("Summary refreshed with %d records.%n", reportRow - 1);
        }
    }

    private static Sheet requireSheet(Workbook workbook, String name) {
        Sheet sheet = workbook.getSheet(name);
        if (sheet == null) throw new IllegalArgumentException("Required worksheet not found: " + name);
        return sheet;
    }

    public static void main(String[] args) throws IOException {
        if (args.length < 1 || args.length > 2) {
            System.err.println("Usage: java vba_sample_1 <workbook.xlsx> [output.xlsx]");
            System.exit(2);
        }
        Path input = Path.of(args[0]);
        Path output = args.length == 2 ? Path.of(args[1]) : VbaExcelSupport.defaultOutput(input);
        RefreshMonthlySummary(input, output);
    }
}
