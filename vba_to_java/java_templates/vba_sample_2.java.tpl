import java.util.Locale;

public final class vba_sample_2 {
    private static final int VB_OBJECT_ERROR = -2147221504;
    public static final Object VBA_EMPTY = new Object();

    public static final class VbaError extends RuntimeException {
        private final int number;
        private final String source;

        VbaError(int number, String source, String description) {
            super(description);
            this.number = number;
            this.source = source;
        }

        public int number() { return number; }
        public String source() { return source; }
        public String description() { return getMessage(); }
    }

    public static String NormalizeSku(Object rawSku) {
        try {
            if (rawSku == null || rawSku == VBA_EMPTY) {
                throw new VbaError(VB_OBJECT_ERROR + 1001, "NormalizeSku", "SKU is required.");
            }
            if (!(rawSku instanceof String)) {
                throw new VbaError(VB_OBJECT_ERROR + 1002, "NormalizeSku", "SKU must be text.");
            }

            String normalized = ((String) rawSku).trim().toUpperCase(Locale.getDefault());
            if (normalized.length() < 4 || normalized.length() > 24) {
                throw new VbaError(VB_OBJECT_ERROR + 1003, "NormalizeSku", "SKU length is invalid.");
            }
            if (!normalized.matches("[A-Z0-9-]+")) {
                throw new VbaError(
                    VB_OBJECT_ERROR + 1004,
                    "NormalizeSku",
                    "SKU contains unsupported characters."
                );
            }
            return normalized;
        } catch (VbaError error) {
            throw new VbaError(error.number(), "NormalizeSku", error.description());
        }
    }

    public static void main(String[] args) {
        if (args.length != 1) {
            System.err.println("Usage: java vba_sample_2 <sku>");
            System.exit(2);
        }
        try {
            System.out.println(NormalizeSku(args[0]));
        } catch (VbaError error) {
            System.err.printf("%d | %s | %s%n", error.number(), error.source(), error.description());
            System.exit(1);
        }
    }
}
