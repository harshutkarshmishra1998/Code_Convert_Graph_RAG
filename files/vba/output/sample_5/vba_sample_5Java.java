import java.util.*;
import java.util.concurrent.TimeUnit;

public class vba_sample_5Java {

    private static final long LOCK_TIMEOUT_SECONDS = 30L;
    private static final int vbRed = 255;
    private static final int vbObjectError = 0x80040000;

    /* Adapter interfaces ---------------------------------------------------- */

    public interface Application {
        boolean getEnableEvents();
        void setEnableEvents(boolean value);
    }

    public interface Workbook {
        Worksheet getWorksheets(String name);
    }

    public interface Worksheet {
        Cell getCell(long row, String column);
        long getRowsCount();
        Range getRange(String start, String end);
        void sort(Range key1, boolean header);
    }

    public interface Cell {
        Object getValue2();
        void setValue2(Object value);
        Interior getInterior();
    }

    public interface Interior {
        void setColor(int color);
    }

    public interface Range {
        // placeholder for range operations
    }

    /* --------------------------------------------------------------------- */

    public static void ReconcileInventory(Worksheet inventorySheet, Worksheet movementSheet,
                                          Application app, Workbook wb) {
        Map<String, Double> inventory = null;
        Map<String, Double> movementTotals = null;
        long lastInventoryRow = 0;
        long lastMovementRow = 0;
        long rowNumber = 0;
        String itemCode = "";
        double quantity = 0.0;
        Date transactionDate = null;
        boolean previousEvents = false;
        double startedAt = 0.0;
        double elapsedSeconds = 0.0;
        long failureNumber = 0;
        String failureDescription = "";

        try {
            previousEvents = app.getEnableEvents();
            app.setEnableEvents(false);
            startedAt = System.currentTimeMillis() / 1000.0;
            inventory = new HashMap<>();
            movementTotals = new HashMap<>();

            lastInventoryRow = inventorySheet.getRowsCount();
            lastMovementRow = movementSheet.getRowsCount();

            for (rowNumber = 2; rowNumber <= lastInventoryRow; rowNumber++) {
                itemCode = ucaseTrimCStr(inventorySheet.getCell(rowNumber, "A").getValue2());
                if (itemCode.length() > 0) {
                    if (isNumeric(inventorySheet.getCell(rowNumber, "B").getValue2())) {
                        if (inventory.containsKey(itemCode)) {
                            raiseError(vbObjectError + 4001, "ReconcileInventory",
                                       "Duplicate inventory item: " + itemCode);
                        }
                        inventory.put(itemCode,
                                      cdbl(inventorySheet.getCell(rowNumber, "B").getValue2()));
                    } else {
                        inventorySheet.getCell(rowNumber, "B").getInterior().setColor(vbRed);
                    }
                }
            }

            for (rowNumber = 2; rowNumber <= lastMovementRow; rowNumber++) {
                itemCode = ucaseTrimCStr(movementSheet.getCell(rowNumber, "A").getValue2());
                if (isDate(movementSheet.getCell(rowNumber, "B").getValue2())
                        && isNumeric(movementSheet.getCell(rowNumber, "C").getValue2())) {
                    transactionDate = cdate(movementSheet.getCell(rowNumber, "B").getValue2());
                    quantity = cdbl(movementSheet.getCell(rowNumber, "C").getValue2());

                    if (dateDiffDays(transactionDate, new Date()) <= 30
                            && !transactionDate.after(new Date())) {
                        if (movementTotals.containsKey(itemCode)) {
                            movementTotals.put(itemCode,
                                    movementTotals.get(itemCode) + quantity);
                        } else {
                            movementTotals.put(itemCode, quantity);
                        }
                    }
                }
            }

            for (rowNumber = 2; rowNumber <= lastInventoryRow; rowNumber++) {
                itemCode = ucaseTrimCStr(inventorySheet.getCell(rowNumber, "A").getValue2());
                if (inventory.containsKey(itemCode)) {
                    quantity = inventory.get(itemCode);
                    if (movementTotals.containsKey(itemCode)) {
                        quantity += movementTotals.get(itemCode);
                    }

                    if (quantity < 0) {
                        inventorySheet.getCell(rowNumber, "C").setValue2("INVESTIGATE");
                    } else if (quantity == 0) {
                        inventorySheet.getCell(rowNumber, "C").setValue2("OUT OF STOCK");
                    } else if (quantity < 5) {
                        inventorySheet.getCell(rowNumber, "C").setValue2("REORDER");
                    } else {
                        inventorySheet.getCell(rowNumber, "C").setValue2("OK");
                    }
                    inventorySheet.getCell(rowNumber, "D").setValue2(quantity);
                }

                elapsedSeconds = System.currentTimeMillis() / 1000.0 - startedAt;
                if (elapsedSeconds < 0) {
                    elapsedSeconds += 86400;
                }
                if (elapsedSeconds > LOCK_TIMEOUT_SECONDS) {
                    raiseError(vbObjectError + 4002, "ReconcileInventory",
                               "Reconciliation exceeded its time limit.");
                }
            }

            inventorySheet.sort(inventorySheet.getRange("A1", "D" + lastInventoryRow), true);

        } catch (VbaException e) {
            failureNumber = e.getNumber();
            failureDescription = e.getDescription();
            LogInventoryError(failureNumber, failureDescription, itemCode, wb);
            System.out.println("Inventory reconciliation failed: " + failureDescription);
        } catch (Exception e) {
            failureNumber = e.hashCode();
            failureDescription = e.getMessage();
            LogInventoryError(failureNumber, failureDescription, itemCode, wb);
            System.out.println("Inventory reconciliation failed: " + failureDescription);
        } finally {
            app.setEnableEvents(previousEvents);
            movementTotals = null;
            inventory = null;
        }
    }

    public static void LogInventoryError(long errorNumber, String description,
                                         String itemCode, Workbook wb) {
        Worksheet logSheet = null;
        long nextLogRow = 0;
        try {
            logSheet = wb.getWorksheets("Error Log");
            if (logSheet == null) {
                return;
            }
            nextLogRow = logSheet.getRowsCount() + 1;
            logSheet.getCell(nextLogRow, "A").setValue2(new Date());
            logSheet.getCell(nextLogRow, "B").setValue2(errorNumber);
            logSheet.getCell(nextLogRow, "C").setValue2(description);
            logSheet.getCell(nextLogRow, "D").setValue2(itemCode);
        } catch (Exception e) {
            // On Error Resume Next: ignore any errors
        }
    }

    /* Helper methods -------------------------------------------------------- */

    private static String ucaseTrimCStr(Object value) {
        return value == null ? "" : value.toString().trim().toUpperCase();
    }

    private static boolean isNumeric(Object value) {
        if (value == null) return false;
        try {
            Double.parseDouble(value.toString());
            return true;
        } catch (NumberFormatException ex) {
            return false;
        }
    }

    private static boolean isDate(Object value) {
        return value instanceof Date;
    }

    private static Date cdate(Object value) {
        return value instanceof Date ? (Date) value : new Date();
    }

    private static double cdbl(Object value) {
        return Double.parseDouble(value.toString());
    }

    private static long dateDiffDays(Date d1, Date d2) {
        long diff = d2.getTime() - d1.getTime();
        return TimeUnit.MILLISECONDS.toDays(diff);
    }

    private static void raiseError(int number, String source, String message) {
        throw new VbaException(number, source, message);
    }

    /* Exception class ------------------------------------------------------- */

    public static class VbaException extends RuntimeException {
        private final int number;
        private final String source;
        private final String description;

        public VbaException(int number, String source, String description) {
            super("Error " + number + " in " + source + ": " + description);
            this.number = number;
            this.source = source;
            this.description = description;
        }

        public int getNumber() {
            return number;
        }

        public String getSource() {
            return source;
        }

        public String getDescription() {
            return description;
        }
    }
}