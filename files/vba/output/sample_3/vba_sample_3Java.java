import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.time.format.DateTimeParseException;
import java.util.HashMap;
import java.util.Map;

public class vba_sample_3Java {

    /* Adapter interfaces ---------------------------------------------------- */

    public interface StreamAdapter {
        void open() throws IOException;
        void loadFromFile(String path) throws IOException;
        boolean isEOS();
        String readText(int count) throws IOException;
        void close() throws IOException;
        int getState(); // 1 = open, 0 = closed
        void setType(int type);
        void setCharset(String charset);
    }

    public interface DictionaryAdapter {
        boolean exists(String key);
        BigDecimal get(String key);
        void put(String key, BigDecimal value);
        void add(String key, BigDecimal value);
    }

    public interface WorkbookAdapter {
        WorksheetAdapter getWorksheet(String name);
    }

    public interface WorksheetAdapter {
        CellAdapter getCell(int row, String column);
        RowsAdapter getRows();
        int getLastUsedRow(String column);
    }

    public interface RowsAdapter {
        int getCount();
    }

    public interface CellAdapter {
        Object getValue();
        void setValue(Object value);
    }

    /* Utility methods ------------------------------------------------------- */

    private static boolean isNumeric(String s) {
        if (s == null) return false;
        try {
            new BigDecimal(s);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private static boolean isDate(Object obj) {
        if (obj instanceof Number) return true;
        if (obj instanceof String) return parseDate((String) obj) != null;
        if (obj instanceof java.util.Date) return true;
        return false;
    }

    private static LocalDate dateValue(Object obj) {
        if (obj instanceof Number) {
            return excelSerialToDate(((Number) obj).doubleValue());
        }
        if (obj instanceof String) {
            return parseDate((String) obj);
        }
        if (obj instanceof java.util.Date) {
            return toLocalDate((java.util.Date) obj);
        }
        return null;
    }

    private static LocalDate toLocalDate(java.util.Date d) {
        return d.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }

    private static LocalDate excelSerialToDate(double serial) {
        LocalDate epoch = LocalDate.of(1899, 12, 30);
        return epoch.plusDays((long) serial);
    }

    private static LocalDate parseDate(String s) {
        if (s == null) return null;
        DateTimeFormatter[] formatters = new DateTimeFormatter[]{
                DateTimeFormatter.ISO_LOCAL_DATE,
                new DateTimeFormatterBuilder()
                        .parseCaseInsensitive()
                        .appendPattern("M/d/yyyy")
                        .toFormatter(),
                new DateTimeFormatterBuilder()
                        .parseCaseInsensitive()
                        .appendPattern("dd-MMM-yyyy")
                        .toFormatter(),
                new DateTimeFormatterBuilder()
                        .parseCaseInsensitive()
                        .appendPattern("dd/MM/yyyy")
                        .toFormatter(),
                new DateTimeFormatterBuilder()
                        .parseCaseInsensitive()
                        .appendPattern("yyyy/MM/dd")
                        .toFormatter()
        };
        for (DateTimeFormatter f : formatters) {
            try {
                return LocalDate.parse(s, f);
            } catch (DateTimeParseException ignored) {
            }
        }
        return null;
    }

    /* Main procedure -------------------------------------------------------- */

    public static void ImportDailySales(String tsvPath, java.util.Date reportDate,
                                        WorkbookAdapter workbook) {
        // Local variables with VBA defaults
        String lineText = "";
        String regionCode = "";
        BigDecimal saleAmount = BigDecimal.ZERO;
        long lastRow = 0;
        long rowNumber = 0;

        StreamAdapter stream = new SimpleStreamAdapter();
        DictionaryAdapter salesByRegion = new SimpleDictionaryAdapter();

        try {
            // File existence check
            File file = new File(tsvPath);
            if (!file.exists()) {
                throw new RuntimeException("ImportDailySales: Input file does not exist.");
            }

            // Initialize stream
            stream.setType(2);
            stream.setCharset("utf-8");
            stream.open();
            stream.loadFromFile(tsvPath);

            if (!stream.isEOS()) {
                lineText = stream.readText(-2);
            }

            while (!stream.isEOS()) {
                lineText = stream.readText(-2);
                String[] fields = lineText.split("\t", -1);

                if (fields.length >= 3) {
                    regionCode = fields[1].trim().toUpperCase();
                    if (isNumeric(fields[2]) && regionCode.length() > 0) {
                        saleAmount = new BigDecimal(fields[2]);

                        if (salesByRegion.exists(regionCode)) {
                            BigDecimal existing = salesByRegion.get(regionCode);
                            salesByRegion.put(regionCode, existing.add(saleAmount));
                        } else {
                            salesByRegion.add(regionCode, saleAmount);
                        }
                    }
                }
            }

            stream.close();
            // Set stream to Nothing equivalent: nothing to do in Java

            WorksheetAdapter ws = workbook.getWorksheet("Daily Sales");
            if (ws == null) {
                throw new RuntimeException("ImportDailySales: Worksheet 'Daily Sales' not found.");
            }

            lastRow = ws.getLastUsedRow("A");

            LocalDate reportLocal = toLocalDate(reportDate);

            for (rowNumber = 2; rowNumber <= lastRow; rowNumber++) {
                CellAdapter cellA = ws.getCell((int) rowNumber, "A");
                Object valA = cellA.getValue();
                if (isDate(valA)) {
                    LocalDate dateA = dateValue(valA);
                    if (dateA != null && dateA.equals(reportLocal)) {
                        CellAdapter cellB = ws.getCell((int) rowNumber, "B");
                        Object valB = cellB.getValue();
                        regionCode = (valB != null) ? valB.toString().trim().toUpperCase() : "";
                        if (salesByRegion.exists(regionCode)) {
                            BigDecimal total = salesByRegion.get(regionCode);
                            ws.getCell((int) rowNumber, "C").setValue(total);
                        } else {
                            ws.getCell((int) rowNumber, "C").setValue(0);
                        }
                    }
                }
            }

        } catch (Exception e) {
            // Error handling similar to VBA Err.Raise
            throw new RuntimeException("ImportDailySales: " + e.getMessage(), e);
        } finally {
            try {
                if (stream != null && stream.getState() == 1) {
                    stream.close();
                }
            } catch (Exception ignored) {
            }
        }
    }

    /* Concrete adapter implementations ------------------------------------- */

    public static class SimpleStreamAdapter implements StreamAdapter {
        private BufferedReader reader;
        private boolean eos = false;
        private int state = 0; // 0 closed, 1 open
        private int type;
        private String charset;

        @Override
        public void open() throws IOException {
            state = 1;
        }

        @Override
        public void loadFromFile(String path) throws IOException {
            FileInputStream fis = new FileInputStream(path);
            InputStreamReader isr = new InputStreamReader(fis, StandardCharsets.UTF_8);
            reader = new BufferedReader(isr);
        }

        @Override
        public boolean isEOS() {
            return eos;
        }

        @Override
        public String readText(int count) throws IOException {
            if (reader == null) return "";
            String line = reader.readLine();
            if (line == null) {
                eos = true;
                return "";
            }
            return line;
        }

        @Override
        public void close() throws IOException {
            if (reader != null) {
                reader.close();
                reader = null;
            }
            state = 0;
        }

        @Override
        public int getState() {
            return state;
        }

        @Override
        public void setType(int type) {
            this.type = type;
        }

        @Override
        public void setCharset(String charset) {
            this.charset = charset;
        }
    }

    public static class SimpleDictionaryAdapter implements DictionaryAdapter {
        private final Map<String, BigDecimal> map = new HashMap<>();

        @Override
        public boolean exists(String key) {
            return map.containsKey(key);
        }

        @Override
        public BigDecimal get(String key) {
            return map.get(key);
        }

        @Override
        public void put(String key, BigDecimal value) {
            map.put(key, value);
        }

        @Override
        public void add(String key, BigDecimal value) {
            map.put(key, value);
        }
    }
}