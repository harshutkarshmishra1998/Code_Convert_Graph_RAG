import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

public final class vba_sample_6 {
    public static final int ModeNone = 0;
    public static final int ModeNormalize = 1;
    public static final int ModeDeduplicate = 2;
    public static final int ModeAggregate = 4;
    public static final int ModeAudit = 8;
    public static final int ModeAll = ModeNormalize | ModeDeduplicate | ModeAggregate | ModeAudit;
    private static final BigDecimal ZERO_CURRENCY = new BigDecimal("0.0000");
    private static final BigDecimal DEDUP_THRESHOLD = new BigDecimal("10000.0000");
    private static final BigDecimal MAX_CURRENCY = new BigDecimal("922337203685477.5807");
    private static final BigDecimal MIN_CURRENCY = new BigDecimal("-922337203685477.5808");

    public static final class LedgerEntry {
        public long EntryId;
        public String AccountCode = "";
        public BigDecimal Amount = ZERO_CURRENCY;
        public LocalDate PostedDate = LocalDate.now();
        public List<String> Tags = new ArrayList<>();
        public boolean IsReconciled;

        LedgerEntry copy() {
            LedgerEntry result = new LedgerEntry();
            result.EntryId = EntryId;
            result.AccountCode = AccountCode;
            result.Amount = Amount;
            result.PostedDate = PostedDate;
            result.Tags = new ArrayList<>(Tags);
            result.IsReconciled = IsReconciled;
            return result;
        }
    }

    public static final class BatchSummary {
        public String BatchId = "";
        public int ModeMask;
        public int RecordCount;
        public BigDecimal TotalAmount = ZERO_CURRENCY;
        public List<LedgerEntry> Entries = new ArrayList<>();
    }

    private static final class InternalGraphNode {
        String NodeKey = "";
        double Weight;
        boolean Visited;
        List<String> Children = new ArrayList<>();
    }

    private static int m_GlobalSequence;
    private static List<String> m_AuditLog;
    private static Map<String, Integer> m_NodeRegistry;

    private vba_sample_6() { }

    private static void InitInfrastructure() {
        if (m_NodeRegistry == null) m_NodeRegistry = new LinkedHashMap<>();
        if (m_AuditLog == null) m_AuditLog = new ArrayList<>();
    }

    private static int NextSequence() {
        return ++m_GlobalSequence;
    }

    private static void LogAudit(String message) {
        if (m_AuditLog == null) m_AuditLog = new ArrayList<>();
        m_AuditLog.add(LocalDateTime.now().format(
            DateTimeFormatter.ofPattern("yyyy-MM-dd hh:mm:ss", Locale.ROOT)
        ) + " | " + message);
    }

    public static String NormalizeLineEndings(String text) {
        return text.replace("\r\n", "\n").replace("\r", "\n");
    }

    public static boolean ExecuteStressPipeline(
        String rawPayload,
        int modeMask,
        StringBuilder outReport,
        int maxIterations
    ) {
        if (outReport != null) outReport.setLength(0);
        if (rawPayload == null || rawPayload.trim().isEmpty()) {
            if (outReport != null) outReport.append("EMPTY_PAYLOAD");
            return false;
        }

        try {
            InitInfrastructure();
            BatchSummary batch = new BatchSummary();
            batch.BatchId = "BATCH_" +
                LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_hhmmss")) +
                "_" + NextSequence();
            batch.ModeMask = modeMask;

            String[] lines = NormalizeLineEndings(rawPayload).split("\n", -1);
            int processedCount = 0;
            for (int idx = 0; idx < lines.length; idx++) {
                if (idx + 1 > maxIterations) {
                    LogAudit("WARN: Max iteration cap reached at row " + idx);
                    break;
                }
                if (!lines[idx].trim().isEmpty()) {
                    LedgerEntry parsed = new LedgerEntry();
                    if (TryParseEntry(lines[idx], idx, parsed)) {
                        batch.Entries.add(parsed);
                        batch.RecordCount++;

                        if ((modeMask & ModeNormalize) == ModeNormalize) {
                            NormalizeEntryInPlace(parsed);
                        }
                        if ((modeMask & ModeAggregate) == ModeAggregate) {
                            batch.TotalAmount = toCurrency(batch.TotalAmount.add(parsed.Amount));
                        }
                        processedCount++;
                    } else {
                        String preview = lines[idx].substring(0, Math.min(60, lines[idx].length()));
                        LogAudit("SKIP: Malformed row " + idx + " -> " + preview);
                    }
                }
                if (idx % 250 == 0) Thread.yield();
            }

            if ((modeMask & ModeDeduplicate) == ModeDeduplicate && batch.RecordCount > 1) {
                DeduplicateBatchEntries(batch);
            }
            String statusMsg;
            if ((modeMask & ModeAudit) == ModeAudit) {
                statusMsg = RunLegacyGoSubAudit(batch);
            } else {
                statusMsg = "AUDIT_SKIPPED";
            }
            String summary = BuildSummaryReport(batch, processedCount, statusMsg);
            if (outReport != null) outReport.append(summary);
            return true;
        } catch (RuntimeException error) {
            String message = "PIPELINE_ERROR [" + error.getClass().getSimpleName() + "]: " +
                error.getMessage() + " at line/step context.";
            if (outReport != null) outReport.append(message);
            LogAudit(message);
            return false;
        }
    }

    public static boolean ExecuteStressPipeline(String rawPayload, int modeMask, StringBuilder outReport) {
        return ExecuteStressPipeline(rawPayload, modeMask, outReport, 1000);
    }

    public static boolean ExecuteStressPipeline(String rawPayload, int modeMask) {
        return ExecuteStressPipeline(rawPayload, modeMask, null, 1000);
    }

    private static boolean TryParseEntry(String rawLine, long fallbackId, LedgerEntry target) {
        try {
            String[] parts = rawLine.split("\\|", -1);
            if (parts.length < 4) return false;
            String id = parts[0].trim();
            try {
                target.EntryId = new BigDecimal(id).setScale(0, RoundingMode.HALF_EVEN).longValueExact();
            } catch (NumberFormatException | ArithmeticException exception) {
                target.EntryId = fallbackId;
            }

            target.AccountCode = parts[1].trim().toUpperCase(Locale.getDefault());
            if (target.AccountCode.isEmpty()) return false;
            double rawAmount = Double.parseDouble(parts[2].trim());
            target.Amount = toCurrency(BigDecimal.valueOf(rawAmount));
            try {
                target.PostedDate = parseDate(parts[3].trim());
            } catch (DateTimeParseException exception) {
                target.PostedDate = LocalDate.now();
            }
            target.IsReconciled = false;

            if (parts.length >= 5 && !parts[4].trim().isEmpty()) {
                String[] tags = parts[4].split(",", -1);
                target.Tags = new ArrayList<>(tags.length);
                for (String tag : tags) target.Tags.add(tag.trim().toLowerCase(Locale.getDefault()));
            } else {
                target.Tags = new ArrayList<>(List.of("untagged"));
            }
            return true;
        } catch (RuntimeException exception) {
            return false;
        }
    }

    private static void NormalizeEntryInPlace(LedgerEntry entry) {
        entry.AccountCode = entry.AccountCode.replace(" ", "_").replace("-", "_");
        entry.IsReconciled = entry.Amount.signum() >= 0 &&
            entry.Amount.compareTo(DEDUP_THRESHOLD) < 0;
        boolean hasCoreTag = false;
        for (int i = 0; i < entry.Tags.size(); i++) {
            String tag = entry.Tags.get(i).trim();
            entry.Tags.set(i, tag);
            if (tag.equalsIgnoreCase("core")) hasCoreTag = true;
        }

        if (!hasCoreTag) entry.Tags.add("core");
    }

    private static BigDecimal toCurrency(BigDecimal amount) {
        BigDecimal rounded = amount.setScale(4, RoundingMode.HALF_EVEN);
        if (rounded.compareTo(MAX_CURRENCY) > 0 || rounded.compareTo(MIN_CURRENCY) < 0) {
            throw new ArithmeticException("Currency value is outside the VBA Currency range");
        }
        return rounded;
    }

    private static void DeduplicateBatchEntries(BatchSummary batch) {
        Map<String, LedgerEntry> seen = new LinkedHashMap<>();
        BigDecimal newTotal = ZERO_CURRENCY;
        for (LedgerEntry entry : batch.Entries) {
            String key = entry.EntryId + "::" + entry.AccountCode;
            if (!seen.containsKey(key)) {
                seen.put(key, entry);
                newTotal = toCurrency(newTotal.add(entry.Amount));
            } else {
                LogAudit("DEDUP: Dropped duplicate key " + key);
            }
        }
        batch.Entries = new ArrayList<>(seen.values());
        batch.RecordCount = batch.Entries.size();
        batch.TotalAmount = newTotal;
    }

    private static String RunLegacyGoSubAudit(BatchSummary batch) {
        if (batch.Entries.isEmpty()) {
            throw new IllegalStateException("Subscript out of range while auditing an empty batch.");
        }
        int flaggedCount = 0;
        for (LedgerEntry entry : batch.Entries) {
            String checkResult;
            if (entry.Amount.compareTo(ZERO_CURRENCY) == 0) {
                checkResult = "ZERO_AMOUNT";
            } else if (entry.PostedDate.getYear() < 1990 || entry.PostedDate.getYear() > 2099) {
                checkResult = "OUT_OF_RANGE_DATE";
            } else {
                checkResult = "OK";
            }
            if (!checkResult.equalsIgnoreCase("OK")) {
                flaggedCount++;
                LogAudit("FLAG [" + entry.EntryId + "]: " + checkResult);
            }
        }
        return "AUDITED_" + batch.RecordCount + "_FLAGGED_" + flaggedCount;
    }

    public static int EvaluateDependencyDepth(String edgesSpec, String startNode) {
        return EvaluateDependencyDepth(edgesSpec, startNode, 0, 25);
    }

    public static int EvaluateDependencyDepth(
        String edgesSpec, String startNode, int currentDepth, int maxDepth
    ) {
        if (currentDepth >= maxDepth) return currentDepth;
        if (edgesSpec == null || edgesSpec.trim().isEmpty() ||
                startNode == null || startNode.trim().isEmpty()) return currentDepth;
        int maxChildDepth = currentDepth;
        for (String pair : edgesSpec.split(",", -1)) {
            String[] edge = pair.split(">", -1);
            if (edge.length != 2) continue;
            String source = edge[0].trim();
            String destination = edge[1].trim();
            if (source.equalsIgnoreCase(startNode.trim()) && !destination.isEmpty() &&
                    !destination.equalsIgnoreCase(source)) {
                int childDepth = EvaluateDependencyDepth(
                    edgesSpec, destination, currentDepth + 1, maxDepth
                );
                maxChildDepth = Math.max(maxChildDepth, childDepth);
            }
        }
        return maxChildDepth;
    }

    public static String BuildUpsertSql(String tableName, LedgerEntry entry) {
        String cleanTable = SanitizeSqlIdentifier(tableName);
        String cleanAcct = entry.AccountCode.replace("'", "''");
        String cleanDate = entry.PostedDate.format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
        int recFlag = entry.IsReconciled ? 1 : 0;
        String amount = entry.Amount.toPlainString().replace(",", ".");
        return "MERGE INTO " + cleanTable + " AS tgt " +
            "USING (SELECT " + entry.EntryId + " AS entry_id, '" + cleanAcct + "' AS account_code) AS src " +
            "ON tgt.entry_id = src.entry_id AND tgt.account_code = src.account_code " +
            "WHEN MATCHED THEN UPDATE SET amount = " + amount +
            ", posted_date = DATE '" + cleanDate + "', is_reconciled = " + recFlag + " " +
            "WHEN NOT MATCHED THEN INSERT (entry_id, account_code, amount, posted_date, is_reconciled) " +
            "VALUES (" + entry.EntryId + ", '" + cleanAcct + "', " + amount +
            ", DATE '" + cleanDate + "', " + recFlag + ");";
    }

    private static String SanitizeSqlIdentifier(String ident) {
        StringBuilder result = new StringBuilder();
        for (int i = 0; i < ident.length(); i++) {
            char ch = ident.charAt(i);
            if ((ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') ||
                    (ch >= '0' && ch <= '9') || ch == '_' || ch == '.') {
                result.append(ch);
            }
        }
        return result.isEmpty() ? "default_ledger_table" : result.toString();
    }

    private static String BuildSummaryReport(
        BatchSummary batch, int parsedCount, String statusMsg
    ) {
        java.text.NumberFormat amountFormat = java.text.NumberFormat.getNumberInstance(Locale.getDefault());
        amountFormat.setMinimumFractionDigits(2);
        amountFormat.setMaximumFractionDigits(2);
        StringBuilder result = new StringBuilder("=== BATCH SUMMARY ===\r\n");
        result.append("Batch ID:      ").append(batch.BatchId).append("\r\n");
        result.append("Mode Mask:     ").append(batch.ModeMask).append("\r\n");
        result.append("Parsed Rows:   ").append(parsedCount).append("\r\n");
        result.append("Final Rows:    ").append(batch.RecordCount).append("\r\n");
        result.append("Total Amount:  ").append(amountFormat.format(batch.TotalAmount)).append("\r\n");
        result.append("Audit Status:  ").append(statusMsg).append("\r\n");
        if (m_AuditLog != null && !m_AuditLog.isEmpty()) {
            result.append("--- RECENT AUDIT LOGS ---\r\n");
            int limit = Math.min(15, m_AuditLog.size());
            for (int i = 0; i < limit; i++) result.append(m_AuditLog.get(i)).append("\r\n");
        }
        return result.toString();
    }

    public static int GetAuditEntryCount() {
        return m_AuditLog == null ? 0 : m_AuditLog.size();
    }

    public static void ResetModuleState() {
        m_GlobalSequence = 0;
        m_AuditLog = new ArrayList<>();
        m_NodeRegistry = null;
    }

    private static LocalDate parseDate(String text) {
        try {
            return LocalDate.parse(text, DateTimeFormatter.ISO_LOCAL_DATE);
        } catch (DateTimeParseException ignored) { }
        for (DateTimeFormatter formatter : new DateTimeFormatter[] {
                DateTimeFormatter.ofPattern("M/d/uuuu"),
                DateTimeFormatter.ofPattern("M/d/uu"),
                DateTimeFormatter.ofPattern("d-MMM-uuuu", Locale.ENGLISH)}) {
            try {
                return LocalDate.parse(text, formatter);
            } catch (DateTimeParseException ignored) { }
        }
        throw new IllegalArgumentException("Invalid entry date: " + text);
    }

    public static void main(String[] args) throws IOException {
        if (args.length < 2 || args.length > 3) {
            System.err.println("Usage: java vba_sample_6 <payload.txt> <mode-mask> [max-iterations]");
            System.exit(2);
        }
        String payload = Files.readString(Path.of(args[0]));
        int modeMask = Integer.parseInt(args[1]);
        int maxIterations = args.length == 3 ? Integer.parseInt(args[2]) : 1000;
        StringBuilder report = new StringBuilder();
        boolean succeeded = ExecuteStressPipeline(payload, modeMask, report, maxIterations);
        System.out.print(report);
        if (!succeeded) System.exit(1);
    }
}
