import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public final class vba_sample_5 {
    private static final long LOCK_TIMEOUT_SECONDS = 30;

    private vba_sample_5() { }

    public static boolean ReconcileInventory(
        Path inputWorkbook, Path outputWorkbook, LocalDateTime now
    ) throws IOException {
        try (Workbook workbook = VbaExcelSupport.open(inputWorkbook)) {
            Sheet inventorySheet = requireSheet(workbook, "Inventory");
            Sheet movementSheet = requireSheet(workbook, "Movements");
            Map<String, Double> inventory = new HashMap<>();
            Map<String, Double> movementTotals = new HashMap<>();
            String itemCode = "";
            boolean succeeded = true;
            long startedAt = System.nanoTime();

            try {
                int lastInventoryRow = VbaExcelSupport.lastRowInColumn(inventorySheet, 0);
                int lastMovementRow = VbaExcelSupport.lastRowInColumn(movementSheet, 0);

                for (int row = 1; row <= lastInventoryRow; row++) {
                    itemCode = code(VbaExcelSupport.cell(inventorySheet, row, 0));
                    if (!itemCode.isEmpty()) {
                        Cell quantityCell = VbaExcelSupport.cell(inventorySheet, row, 1);
                        if (VbaExcelSupport.isNumeric(quantityCell)) {
                            if (inventory.containsKey(itemCode)) {
                                throw new VbaFailure(-2147217503,
                                    "Duplicate inventory item: " + itemCode);
                            }
                            inventory.put(itemCode, VbaExcelSupport.decimal(quantityCell).doubleValue());
                        } else {
                            VbaExcelSupport.fill(
                                VbaExcelSupport.writableCell(inventorySheet, row, 1),
                                org.apache.poi.ss.usermodel.IndexedColors.RED
                            );
                        }
                    }
                }

                for (int row = 1; row <= lastMovementRow; row++) {
                    itemCode = code(VbaExcelSupport.cell(movementSheet, row, 0));
                    Cell dateCell = VbaExcelSupport.cell(movementSheet, row, 1);
                    Cell quantityCell = VbaExcelSupport.cell(movementSheet, row, 2);
                    if (VbaExcelSupport.isDate(dateCell) && VbaExcelSupport.isNumeric(quantityCell)) {
                        LocalDateTime transactionDate = VbaExcelSupport.dateTime(dateCell);
                        double quantity = VbaExcelSupport.decimal(quantityCell).doubleValue();
                        long daysOld = ChronoUnit.DAYS.between(
                            transactionDate.toLocalDate(), now.toLocalDate()
                        );
                        if (!transactionDate.isAfter(now) && daysOld <= 30) {
                            movementTotals.merge(itemCode, quantity, Double::sum);
                        }
                    }
                }

                for (int row = 1; row <= lastInventoryRow; row++) {
                    itemCode = code(VbaExcelSupport.cell(inventorySheet, row, 0));
                    if (inventory.containsKey(itemCode)) {
                        double quantity = inventory.get(itemCode) +
                            movementTotals.getOrDefault(itemCode, 0.0);
                        String status = quantity < 0 ? "INVESTIGATE"
                            : quantity == 0 ? "OUT OF STOCK"
                            : quantity < 5 ? "REORDER" : "OK";
                        VbaExcelSupport.set(
                            VbaExcelSupport.writableCell(inventorySheet, row, 2), status
                        );
                        VbaExcelSupport.set(
                            VbaExcelSupport.writableCell(inventorySheet, row, 3), quantity
                        );
                    }
                    if (System.nanoTime() - startedAt > LOCK_TIMEOUT_SECONDS * 1_000_000_000L) {
                        throw new VbaFailure(-2147217502,
                            "Reconciliation exceeded its time limit.");
                    }
                }

                sortInventory(inventorySheet, lastInventoryRow);
            } catch (RuntimeException error) {
                int errorNumber = error instanceof VbaFailure failure
                    ? failure.number : error instanceof IllegalArgumentException ? 13 : 5;
                appendError(workbook, now, errorNumber, error.getMessage(), itemCode);
                System.err.println("Inventory reconciliation failed: " + error.getMessage());
                succeeded = false;
            }

            VbaExcelSupport.writeAtomically(workbook, outputWorkbook);
            if (succeeded) System.out.println("Inventory reconciliation completed.");
            return succeeded;
        }
    }

    private static String code(Cell cell) {
        return VbaExcelSupport.text(cell).trim().toUpperCase(Locale.getDefault());
    }

    private static Sheet requireSheet(Workbook workbook, String name) {
        Sheet sheet = workbook.getSheet(name);
        if (sheet == null) throw new IllegalArgumentException("Required worksheet not found: " + name);
        return sheet;
    }

    private static void appendError(
        Workbook workbook, LocalDateTime now, int number, String description, String itemCode
    ) {
        Sheet log = workbook.getSheet("Error Log");
        if (log == null) {
            System.err.println("Error Log worksheet is missing; error details were not saved.");
            return;
        }
        int rowIndex = Math.max(1, VbaExcelSupport.lastRowInColumn(log, 0) + 1);
        VbaExcelSupport.set(VbaExcelSupport.writableCell(log, rowIndex, 0), now);
        VbaExcelSupport.set(VbaExcelSupport.writableCell(log, rowIndex, 1), number);
        VbaExcelSupport.set(
            VbaExcelSupport.writableCell(log, rowIndex, 2),
            description == null ? "Unspecified reconciliation error" : description
        );
        VbaExcelSupport.set(VbaExcelSupport.writableCell(log, rowIndex, 3), itemCode);
    }

    private static void sortInventory(Sheet sheet, int lastRow) {
        if (lastRow <= 1) return;
        List<InventoryRow> rows = new ArrayList<>();
        for (int row = 1; row <= lastRow; row++) {
            Object[] values = new Object[4];
            CellStyle[] styles = new CellStyle[4];
            for (int column = 0; column < 4; column++) {
                Cell cell = VbaExcelSupport.cell(sheet, row, column);
                values[column] = VbaExcelSupport.value(cell);
                styles[column] = cell == null ? null : cell.getCellStyle();
            }
            rows.add(new InventoryRow(values, styles));
        }
        rows.sort(Comparator.comparing(
            row -> row.values[0] == null ? "" : row.values[0].toString(),
            String.CASE_INSENSITIVE_ORDER
        ));
        for (int row = 1; row <= lastRow; row++) {
            InventoryRow values = rows.get(row - 1);
            for (int column = 0; column < 4; column++) {
                Cell cell = VbaExcelSupport.writableCell(sheet, row, column);
                VbaExcelSupport.set(cell, values.values[column]);
                if (values.styles[column] != null) cell.setCellStyle(values.styles[column]);
            }
        }
    }

    private record InventoryRow(Object[] values, CellStyle[] styles) { }

    private static final class VbaFailure extends RuntimeException {
        final int number;
        VbaFailure(int number, String message) {
            super(message);
            this.number = number;
        }
    }

    public static void main(String[] args) throws IOException {
        if (args.length < 1 || args.length > 2) {
            System.err.println("Usage: java vba_sample_5 <inventory-workbook.xlsx> [output.xlsx]");
            System.exit(2);
        }
        Path input = Path.of(args[0]);
        Path output = args.length == 2 ? Path.of(args[1]) : VbaExcelSupport.defaultOutput(input);
        boolean succeeded = ReconcileInventory(input, output, LocalDateTime.now());
        if (!succeeded) System.exit(1);
    }
}
