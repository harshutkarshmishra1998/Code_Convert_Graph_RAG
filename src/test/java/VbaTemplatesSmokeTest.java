import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Comparator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class VbaTemplatesSmokeTest {
    private VbaTemplatesSmokeTest() { }

    public static void main(String[] args) throws Exception {
        Path root = Files.createTempDirectory("vba-java-smoke-");
        try {
            testSample1(root);
            testSample2();
            testSample3(root);
            testSample4(root);
            testSample5(root);
            testSample5Duplicate(root);
            testSample6(root);
            System.out.println("All six VBA-to-Java smoke tests passed.");
        } finally {
            try (var paths = Files.walk(root)) {
                paths.sorted(Comparator.reverseOrder()).forEach(path -> {
                    try {
                        Files.deleteIfExists(path);
                    } catch (IOException exception) {
                        throw new RuntimeException("Unable to clean test fixture " + path, exception);
                    }
                });
            }
        }
    }

    private static void testSample1(Path root) throws Exception {
        Path input = root.resolve("sample1-input.xlsx");
        Path output = root.resolve("sample1-output.xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet transactions = workbook.createSheet("Transactions");
            transactions.createRow(0).createCell(0).setCellValue("Account");
            transactions.getRow(0).createCell(1).setCellValue("Date");
            transactions.getRow(0).createCell(2).setCellValue("Amount");
            transactions.createRow(1).createCell(0).setCellValue("acct-1");
            transactions.getRow(1).createCell(1).setCellValue(LocalDateTime.of(2024, 3, 20, 0, 0));
            transactions.getRow(1).createCell(2).setCellValue(42);
            transactions.createRow(2).createCell(0).setCellValue("acct-2");
            transactions.getRow(2).createCell(1).setCellValue("not a date");
            transactions.getRow(2).createCell(2).setCellValue(10);
            workbook.createSheet("Monthly Summary").createRow(0).createCell(0).setCellValue("Account");
            write(workbook, input);
        }

        vba_sample_1.RefreshMonthlySummary(input, output);
        try (Workbook result = WorkbookFactory.create(output.toFile())) {
            Sheet source = result.getSheet("Transactions");
            Sheet report = result.getSheet("Monthly Summary");
            check("acct-1".equals(report.getRow(1).getCell(0).getStringCellValue()),
                "sample 1 should copy the valid account");
            check(report.getRow(1).getCell(2).getNumericCellValue() == 42,
                "sample 1 should copy the transaction amount");
            check(source.getRow(2).getCell(1).getCellStyle().getFillForegroundColor() ==
                IndexedColors.YELLOW.getIndex(), "sample 1 should highlight invalid dates");
        }
    }

    private static void testSample2() {
        check("AB-123".equals(vba_sample_2.NormalizeSku(" ab-123 ")),
            "sample 2 should trim and normalize SKUs");
        boolean rejected = false;
        try {
            vba_sample_2.NormalizeSku("bad!");
        } catch (vba_sample_2.VbaError expected) {
            rejected = expected.number() == -2147220500;
        }
        check(rejected, "sample 2 should preserve its invalid-character error number");
    }

    private static void testSample3(Path root) throws Exception {
        Path input = root.resolve("sample3-input.xlsx");
        Path output = root.resolve("sample3-output.xlsx");
        Path tsv = root.resolve("sales.tsv");
        Files.writeString(tsv, "id\tregion\tamount\n1\tEAST\t12.34\n2\tEAST\t2.66\n3\tSOUTH\t-1.25\n");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Daily Sales");
            sheet.createRow(0).createCell(0).setCellValue("Date");
            sheet.getRow(0).createCell(1).setCellValue("Region");
            sheet.getRow(0).createCell(2).setCellValue("Sales");
            sheet.createRow(1).createCell(0).setCellValue(LocalDateTime.of(2024, 3, 20, 0, 0));
            sheet.getRow(1).createCell(1).setCellValue("east");
            sheet.createRow(2).createCell(0).setCellValue(LocalDateTime.of(2024, 3, 20, 0, 0));
            sheet.getRow(2).createCell(1).setCellValue("NORTH");
            write(workbook, input);
        }

        vba_sample_3.ImportDailySales(
            input, tsv, LocalDate.of(2024, 3, 20), output
        );
        try (Workbook result = WorkbookFactory.create(output.toFile())) {
            Sheet sheet = result.getSheet("Daily Sales");
            check(sheet.getRow(1).getCell(2).getNumericCellValue() == 15,
                "sample 3 should aggregate amounts by normalized region");
            check(sheet.getRow(2).getCell(2).getNumericCellValue() == 0,
                "sample 3 should write zero for a missing region");
        }
    }

    private static void testSample4(Path root) throws Exception {
        Path folder = Files.createDirectory(root.resolve("sample4-input"));
        createDepartmentWorkbook(folder.resolve("a.xlsx"), "A", "A-1");
        createDepartmentWorkbook(folder.resolve("b.xlsx"), "B", "B-1");
        Path output = folder.resolve("consolidated.xlsx");
        check(vba_sample_4.ConsolidateDepartmentFiles(folder, output) == 2,
            "sample 4 should count both source workbooks");
        try (Workbook result = WorkbookFactory.create(output.toFile())) {
            Sheet sheet = result.getSheet("Consolidated");
            check("Department".equals(sheet.getRow(0).getCell(0).getStringCellValue()),
                "sample 4 should retain the first workbook's header");
            check("A-1".equals(sheet.getRow(1).getCell(0).getStringCellValue()),
                "sample 4 should copy first workbook data");
            check("B-1".equals(sheet.getRow(2).getCell(0).getStringCellValue()),
                "sample 4 should skip later workbook headers");
            check(sheet.getLastRowNum() == 2, "sample 4 should append two data rows");
        }
    }

    private static void testSample5(Path root) throws Exception {
        Path input = root.resolve("sample5-input.xlsx");
        Path output = root.resolve("sample5-output.xlsx");
        LocalDateTime now = LocalDateTime.of(2024, 3, 30, 12, 0);
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet inventory = workbook.createSheet("Inventory");
            inventory.createRow(0).createCell(0).setCellValue("Item");
            inventory.getRow(0).createCell(1).setCellValue("Quantity");
            inventory.getRow(0).createCell(2).setCellValue("Status");
            inventory.getRow(0).createCell(3).setCellValue("New quantity");
            inventory.createRow(1).createCell(0).setCellValue("C");
            inventory.getRow(1).createCell(1).setCellValue("bad quantity");
            inventory.createRow(2).createCell(0).setCellValue("B");
            inventory.getRow(2).createCell(1).setCellValue(0);
            inventory.createRow(3).createCell(0).setCellValue("A");
            inventory.getRow(3).createCell(1).setCellValue(2);
            Sheet movements = workbook.createSheet("Movements");
            movements.createRow(0).createCell(0).setCellValue("Item");
            movements.getRow(0).createCell(1).setCellValue("Date");
            movements.getRow(0).createCell(2).setCellValue("Quantity");
            movements.createRow(1).createCell(0).setCellValue("A");
            movements.getRow(1).createCell(1).setCellValue(LocalDateTime.of(2024, 3, 20, 12, 0));
            movements.getRow(1).createCell(2).setCellValue(4);
            workbook.createSheet("Error Log");
            write(workbook, input);
        }

        check(vba_sample_5.ReconcileInventory(input, output, now),
            "sample 5 should reconcile a valid workbook");
        try (Workbook result = WorkbookFactory.create(output.toFile())) {
            Sheet inventory = result.getSheet("Inventory");
            check("A".equals(inventory.getRow(1).getCell(0).getStringCellValue()),
                "sample 5 should sort inventory by item code");
            check("OK".equals(inventory.getRow(1).getCell(2).getStringCellValue()),
                "sample 5 should classify updated positive stock");
            check(inventory.getRow(1).getCell(3).getNumericCellValue() == 6,
                "sample 5 should apply in-window movements");
            check("OUT OF STOCK".equals(inventory.getRow(2).getCell(2).getStringCellValue()),
                "sample 5 should classify zero stock");
            check(inventory.getRow(3).getCell(1).getCellStyle().getFillForegroundColor() ==
                IndexedColors.RED.getIndex(), "sample 5 should highlight invalid quantities");
        }
    }

    private static void testSample6(Path root) throws Exception {
        String payload = "1|a|10|2024-01-01|x\n" +
            "2|b|5|2099-12-31|core\n" +
            "1|a|10|2024-01-01|y";
        StringBuilder report = new StringBuilder();
        check(vba_sample_6.ExecuteStressPipeline(
            payload, vba_sample_6.ModeAll, report, 1000
        ), "sample 6 should accept valid payload rows");
        check(report.toString().contains("Parsed Rows:   3"),
            "sample 6 should report all parsed input rows");
        check(report.toString().contains("Final Rows:    2"),
            "sample 6 should deduplicate records");
        check(report.toString().contains("Total Amount:  15.00"),
            "sample 6 should aggregate the deduplicated currency total");
        check(vba_sample_6.EvaluateDependencyDepth("A>B,B>C,A>D", "A") == 2,
            "sample 6 should compute recursive dependency depth");
        check(vba_sample_6.BuildUpsertSql("ledger; DROP", new vba_sample_6.LedgerEntry())
            .contains("ledgerDROP"), "sample 6 should sanitize SQL identifiers");
        StringBuilder emptyAuditReport = new StringBuilder();
        check(!vba_sample_6.ExecuteStressPipeline(
            "malformed", vba_sample_6.ModeAudit, emptyAuditReport, 1000
        ), "sample 6 should preserve the VBA empty-array audit failure");
    }

    private static void testSample5Duplicate(Path root) throws Exception {
        Path input = root.resolve("sample5-duplicate-input.xlsx");
        Path output = root.resolve("sample5-duplicate-output.xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet inventory = workbook.createSheet("Inventory");
            inventory.createRow(0).createCell(0).setCellValue("Item");
            inventory.getRow(0).createCell(1).setCellValue("Quantity");
            inventory.createRow(1).createCell(0).setCellValue("D");
            inventory.getRow(1).createCell(1).setCellValue(1);
            inventory.createRow(2).createCell(0).setCellValue("D");
            inventory.getRow(2).createCell(1).setCellValue(2);
            Sheet movements = workbook.createSheet("Movements");
            movements.createRow(0).createCell(0).setCellValue("Item");
            movements.getRow(0).createCell(1).setCellValue("Date");
            movements.getRow(0).createCell(2).setCellValue("Quantity");
            Sheet log = workbook.createSheet("Error Log");
            log.createRow(0).createCell(0).setCellValue("Timestamp");
            write(workbook, input);
        }

        check(!vba_sample_5.ReconcileInventory(
            input, output, LocalDateTime.of(2024, 3, 30, 12, 0)
        ), "sample 5 should report duplicate inventory keys as a failure");
        try (Workbook result = WorkbookFactory.create(output.toFile())) {
            Sheet log = result.getSheet("Error Log");
            check(log.getRow(1).getCell(1).getNumericCellValue() == -2147217503,
                "sample 5 should preserve the duplicate-key VBA error number");
            check(log.getRow(1).getCell(3).getStringCellValue().equals("D"),
                "sample 5 should include the duplicate item code in its error log");
        }
    }

    private static void createDepartmentWorkbook(Path path, String header, String value)
        throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Department");
            sheet.createRow(0).createCell(0).setCellValue("Department");
            sheet.getRow(0).createCell(1).setCellValue(header);
            sheet.createRow(1).createCell(0).setCellValue(value);
            write(workbook, path);
        }
    }

    private static void write(Workbook workbook, Path path) throws IOException {
        try (var stream = Files.newOutputStream(path)) {
            workbook.write(stream);
        }
    }

    private static void check(boolean condition, String message) {
        if (!condition) throw new AssertionError(message);
    }
}
