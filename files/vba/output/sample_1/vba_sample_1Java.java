import java.util.Date;
import java.util.Locale;
import java.text.SimpleDateFormat;
import java.text.ParseException;

public class vba_sample_1Java {

    /* Adapter interfaces ---------------------------------------------------- */

    public interface WorkbookAdapter {
        WorksheetAdapter getWorksheet(String name);
    }

    public interface WorksheetAdapter {
        CellAdapter getCell(int row, String column);
        int getRowsCount();
        void autoFitColumns(String columns);
    }

    public interface CellAdapter {
        Object getValue();
        void setValue(Object value);
        InteriorAdapter getInterior();
        CellAdapter end(XlDirection direction);
    }

    public interface InteriorAdapter {
        void setColor(int color);
    }

    public interface MsgBoxAdapter {
        void show(String message, int icon);
    }

    public enum XlDirection {
        xlUp
    }

    /* Constants -------------------------------------------------------------- */

    public static final int vbYellow = 65535;
    public static final int vbInformation = 1;
    public static final int vbExclamation = 16;

    /* Main procedure ---------------------------------------------------------- */

    public void RefreshMonthlySummary(WorkbookAdapter workbook, MsgBoxAdapter msgBox) {
        WorksheetAdapter sourceSheet = null;
        WorksheetAdapter reportSheet = null;
        long lastRow = 0;
        long reportRow = 0;

        try {
            sourceSheet = workbook.getWorksheet("Transactions");
            reportSheet = workbook.getWorksheet("Monthly Summary");
            lastRow = sourceSheet.getRowsCount();
            // Find last used row in column A
            CellAdapter lastCell = sourceSheet.getCell((int) lastRow, "A").end(XlDirection.xlUp);
            lastRow = (long) lastCell.getValue();

            reportRow = 2;

            for (long rowNumber = 2; rowNumber <= lastRow; rowNumber++) {
                CellAdapter idCell = sourceSheet.getCell((int) rowNumber, "A");
                String idValue = toString(idCell.getValue()).trim();
                if (idValue.length() > 0) {
                    CellAdapter dateCell = sourceSheet.getCell((int) rowNumber, "B");
                    Object dateObj = dateCell.getValue();
                    if (isDate(dateObj)) {
                        reportSheet.getCell((int) reportRow, "A").setValue(idCell.getValue());
                        reportSheet.getCell((int) reportRow, "B").setValue(toDate(dateObj));
                        reportSheet.getCell((int) reportRow, "C").setValue(sourceSheet.getCell((int) rowNumber, "C").getValue());
                        reportRow = reportRow + 1;
                    } else {
                        dateCell.getInterior().setColor(vbYellow);
                    }
                }
            }

            reportSheet.autoFitColumns("A:C");
            msgBox.show("Summary refreshed with " + (reportRow - 2) + " records.", vbInformation);
            return;

        } catch (Exception e) {
            msgBox.show("Unable to refresh the monthly summary: " + e.getMessage(), vbExclamation);
        }
    }

    /* Helper methods ---------------------------------------------------------- */

    private boolean isDate(Object obj) {
        if (obj == null) return false;
        if (obj instanceof Date) return true;
        if (obj instanceof Number) return true; // Excel dates are numeric
        String s = toString(obj);
        return parseDate(s) != null;
    }

    private Date toDate(Object obj) {
        if (obj instanceof Date) return (Date) obj;
        if (obj instanceof Number) {
            // Excel serial date to Java Date (simplified)
            double serial = ((Number) obj).doubleValue();
            return excelSerialToDate(serial);
        }
        String s = toString(obj);
        Date d = parseDate(s);
        return d != null ? d : new Date();
    }

    private String toString(Object obj) {
        return obj == null ? "" : obj.toString();
    }

    private Date parseDate(String s) {
        try {
            SimpleDateFormat df = new SimpleDateFormat("M/d/yyyy", Locale.US);
            return df.parse(s);
        } catch (ParseException e) {
            return null;
        }
    }

    private Date excelSerialToDate(double serial) {
        // Excel's day 1 is 1900-01-01; this is a simplified conversion
        long epoch = 25569L; // days between 1900-01-01 and 1970-01-01
        long days = (long) serial;
        long millis = (long) ((serial - days) * 86400000L);
        return new Date((days - epoch) * 86400000L + millis);
    }
}