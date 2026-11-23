import java.util.regex.Pattern;

public class vba_sample_2Java {

    public interface ExcelAdapter {
        void raiseError(int code, String source, String description) throws VbaException;
    }

    public static class VbaException extends Exception {
        public final int code;
        public final String source;
        public final String description;

        public VbaException(int code, String source, String description) {
            super(description);
            this.code = code;
            this.source = source;
            this.description = description;
        }
    }

    public static final int vbObjectError = -2147221504; // typical VBA vbObjectError value
    public static final int vbString = 8; // VarType for String

    public static String NormalizeSku(Object rawSku, ExcelAdapter adapter) throws VbaException {
        String normalized = "";
        try {
            if (rawSku == null || (rawSku instanceof String && ((String) rawSku).isEmpty())) {
                adapter.raiseError(vbObjectError + 1001, "NormalizeSku", "SKU is required.");
            }

            if (VarType(rawSku) != vbString) {
                adapter.raiseError(vbObjectError + 1002, "NormalizeSku", "SKU must be text.");
            }

            normalized = UCase$(Trim$(CStr(rawSku)));

            if (Len(normalized) < 4 || Len(normalized) > 24) {
                adapter.raiseError(vbObjectError + 1003, "NormalizeSku", "SKU length is invalid.");
            }

            if (Like(normalized, "*[!A-Z0-9-]*")) {
                adapter.raiseError(vbObjectError + 1004, "NormalizeSku", "SKU contains unsupported characters.");
            }

            return normalized;
        } catch (VbaException e) {
            adapter.raiseError(e.code, e.source, e.description);
            throw e; // rethrow for caller
        } catch (Exception e) {
            adapter.raiseError(vbObjectError + 1000, "NormalizeSku", e.getMessage());
            throw new VbaException(vbObjectError + 1000, "NormalizeSku", e.getMessage());
        }
    }

    private static int VarType(Object obj) {
        if (obj instanceof String) return vbString;
        return 0;
    }

    private static String CStr(Object obj) {
        return obj == null ? "" : obj.toString();
    }

    private static String Trim$(String s) {
        return s.trim();
    }

    private static String UCase$(String s) {
        return s.toUpperCase();
    }

    private static int Len(String s) {
        return s.length();
    }

    private static boolean Like(String text, String pattern) {
        // Convert VBA Like pattern to regex
        String regex = pattern
                .replace("*", ".*")
                .replace("?", ".")
                .replace("[!A-Z0-9-]*", ".*[^A-Z0-9-].*");
        return Pattern.compile(regex, Pattern.DOTALL).matcher(text).matches();
    }
}