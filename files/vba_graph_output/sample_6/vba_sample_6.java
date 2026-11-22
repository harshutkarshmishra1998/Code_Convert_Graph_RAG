import java.lang.reflect.Array;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

public class vba_sample_6 {
    // Generated from normalized VBA IR. Host/API adapters and VBA compatibility
    // behavior listed in java_codegen_report.json still require implementation.
    private static final Object ModeNone = 0;
    private static final Object ModeNormalize = 1;
    private static final Object ModeDeduplicate = 2;
    private static final Object ModeAggregate = 4;
    private static final Object ModeAudit = 8;
    private static final Object ModeAll = VbaRuntime.binary("or", VbaRuntime.binary("or", VbaRuntime.binary("or", ModeNormalize, ModeDeduplicate), ModeAggregate), ModeAudit);
    // VBA declaration retained in IR: LedgerEntry
    // VBA declaration retained in IR: BatchSummary
    // VBA declaration retained in IR: InternalGraphNode
    private static Object m_GlobalSequence = null;
    private static Object m_AuditLog = null;
    private static Object m_NodeRegistry = null;
    public Object ExecuteStressPipeline(Object rawPayload, Object modeMask, Object outReport, Object maxIterations) {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        Object lines = null;
        Object batch = null;
        Object idx = null;
        Object processedCount = null;
        Object tempAmount = null;
        Object statusMsg = null;
        Object ExecuteStressPipeline = null;
        VbaRuntime.onError("goto_handler", "ErrHandler");
        if (VbaRuntime.truthy(VbaRuntime.binary("=", VbaRuntime.builtin("Len", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{rawPayload})}), 0))) {
            outReport = "EMPTY_PAYLOAD";
            ExecuteStressPipeline = Boolean.FALSE;
            if (VbaRuntime.returnNow()) return ExecuteStressPipeline;
        }
        this.InitInfrastructure();
        VbaRuntime.set(batch, "BatchId", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", "BATCH_", VbaRuntime.builtin("Format$", new Object[]{VbaRuntime.name("Now"), "yyyymmdd_hhnnss"})), "_"), VbaRuntime.builtin("CStr", new Object[]{this.NextSequence()})));
        VbaRuntime.set(batch, "ModeMask", modeMask);
        VbaRuntime.set(batch, "RecordCount", 0);
        VbaRuntime.set(batch, "TotalAmount", "0");
        lines = VbaRuntime.builtin("Split", new Object[]{this.NormalizeLineEndings(rawPayload), VbaRuntime.name("vbLf")});
        VbaRuntime.unsupportedControlFlow("ReDim target");
        idx = VbaRuntime.builtin("LBound", new Object[]{lines});
        for (; VbaRuntime.forCondition(idx, VbaRuntime.builtin("UBound", new Object[]{lines}), 1); idx = VbaRuntime.binary("+", idx, 1)) {
            if (VbaRuntime.truthy(VbaRuntime.binary(">", VbaRuntime.binary("+", VbaRuntime.binary("-", idx, VbaRuntime.builtin("LBound", new Object[]{lines})), 1), maxIterations))) {
                this.LogAudit(VbaRuntime.binary("&", "WARN: Max iteration cap reached at row ", VbaRuntime.builtin("CStr", new Object[]{idx})));
                if (VbaRuntime.returnNow()) break;
            }
            if (VbaRuntime.truthy(VbaRuntime.binary(">", VbaRuntime.builtin("Len", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.index(lines, new Object[]{idx})})}), 0))) {
                if (VbaRuntime.truthy(this.TryParseEntry(VbaRuntime.index(lines, new Object[]{idx}), idx, VbaRuntime.index(VbaRuntime.get(batch, "Entries" ), new Object[]{VbaRuntime.binary("+", VbaRuntime.get(batch, "RecordCount" ), 1)})))) {
                    VbaRuntime.set(batch, "RecordCount", VbaRuntime.binary("+", VbaRuntime.get(batch, "RecordCount" ), 1));
                    if (VbaRuntime.truthy(VbaRuntime.binary("=", (VbaRuntime.binary("and", modeMask, ModeNormalize)), ModeNormalize))) {
                        this.NormalizeEntryInPlace(VbaRuntime.index(VbaRuntime.get(batch, "Entries" ), new Object[]{VbaRuntime.get(batch, "RecordCount" )}));
                    }
                    tempAmount = VbaRuntime.get(VbaRuntime.index(VbaRuntime.get(batch, "Entries" ), new Object[]{VbaRuntime.get(batch, "RecordCount" )}), "Amount" );
                    if (VbaRuntime.truthy(VbaRuntime.binary("=", (VbaRuntime.binary("and", modeMask, ModeAggregate)), ModeAggregate))) {
                        VbaRuntime.set(batch, "TotalAmount", VbaRuntime.binary("+", VbaRuntime.get(batch, "TotalAmount" ), tempAmount));
                    }
                    processedCount = VbaRuntime.binary("+", processedCount, 1);
                } else {
                    this.LogAudit(VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", "SKIP: Malformed row ", VbaRuntime.builtin("CStr", new Object[]{idx})), " -> "), VbaRuntime.builtin("Left$", new Object[]{VbaRuntime.index(lines, new Object[]{idx}), 60})));
                }
            }
            if (VbaRuntime.truthy(VbaRuntime.binary("=", (VbaRuntime.binary("mod", idx, 250)), 0))) {
                VbaRuntime.builtin("DoEvents", new Object[]{});
            }
        }
        if (VbaRuntime.truthy(VbaRuntime.binary(">", VbaRuntime.get(batch, "RecordCount" ), 0))) {
            VbaRuntime.unsupportedControlFlow("ReDim target");
        } else {
            VbaRuntime.unsupportedControlFlow("Erase target");
        }
        if (VbaRuntime.truthy(VbaRuntime.binary("and", VbaRuntime.binary("=", (VbaRuntime.binary("and", modeMask, ModeDeduplicate)), ModeDeduplicate), VbaRuntime.binary(">", VbaRuntime.get(batch, "RecordCount" ), 1)))) {
            this.DeduplicateBatchEntries(batch);
        }
        if (VbaRuntime.truthy(VbaRuntime.binary("=", (VbaRuntime.binary("and", modeMask, ModeAudit)), ModeAudit))) {
            this.RunLegacyGoSubAudit(batch, statusMsg);
        } else {
            statusMsg = "AUDIT_SKIPPED";
        }
        outReport = this.BuildSummaryReport(batch, processedCount, statusMsg);
        ExecuteStressPipeline = Boolean.TRUE;
        if (VbaRuntime.returnNow()) return ExecuteStressPipeline;
        // VBA label: ErrHandler
        outReport = VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", "PIPELINE_ERROR [", VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.get(VbaRuntime.errorState(), "Number" )})), "]: "), VbaRuntime.get(VbaRuntime.errorState(), "Description" )), " at line/step context.");
        this.LogAudit(outReport);
        ExecuteStressPipeline = Boolean.FALSE;
        return ExecuteStressPipeline;
    }

    private void InitInfrastructure() {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        VbaRuntime.onError("resume_next", "");
        m_NodeRegistry = VbaRuntime.builtin("CreateObject", new Object[]{"Scripting.Dictionary"});
        if (VbaRuntime.truthy(VbaRuntime.binary("or", VbaRuntime.binary("<>", VbaRuntime.get(VbaRuntime.errorState(), "Number" ), 0), VbaRuntime.binary("is", m_NodeRegistry, null)))) {
            VbaRuntime.call(VbaRuntime.errorState(), "Clear", new Object[]{});
            m_NodeRegistry = VbaRuntime.newObject("Collection");
        }
        VbaRuntime.onError("disable_handler", "");
        if (VbaRuntime.truthy(VbaRuntime.binary("is", m_AuditLog, null))) {
            m_AuditLog = VbaRuntime.newObject("Collection");
        }
    }

    private Object NextSequence() {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        Object NextSequence = null;
        m_GlobalSequence = VbaRuntime.binary("+", m_GlobalSequence, 1);
        NextSequence = m_GlobalSequence;
        return NextSequence;
    }

    private void LogAudit(Object message) {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        if (VbaRuntime.truthy(VbaRuntime.binary("is", m_AuditLog, null))) {
            m_AuditLog = VbaRuntime.newObject("Collection");
        }
        VbaRuntime.call(m_AuditLog, "Add", new Object[]{VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.builtin("Format$", new Object[]{VbaRuntime.name("Now"), "yyyy-mm-dd hh:nn:ss"}), " | "), message)});
    }

    public Object NormalizeLineEndings(Object text) {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        Object result = null;
        Object NormalizeLineEndings = null;
        result = VbaRuntime.builtin("Replace", new Object[]{text, VbaRuntime.name("vbCrLf"), VbaRuntime.name("vbLf")});
        result = VbaRuntime.builtin("Replace", new Object[]{result, VbaRuntime.name("vbCr"), VbaRuntime.name("vbLf")});
        NormalizeLineEndings = result;
        return NormalizeLineEndings;
    }

    private Object TryParseEntry(Object rawLine, Object fallbackId, Object target) {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        Object parts = null;
        Object tagTokens = null;
        Object i = null;
        Object rawAmt = null;
        Object TryParseEntry = null;
        VbaRuntime.onError("goto_handler", "ParseFail");
        parts = VbaRuntime.builtin("Split", new Object[]{rawLine, "|"});
        if (VbaRuntime.truthy(VbaRuntime.unary("not", VbaRuntime.builtin("IsArray", new Object[]{parts})))) {
            VbaRuntime.unsupportedControlFlow("goto ParseFail");
        }
        if (VbaRuntime.truthy(VbaRuntime.binary("<", VbaRuntime.binary("+", VbaRuntime.binary("-", VbaRuntime.builtin("UBound", new Object[]{parts}), VbaRuntime.builtin("LBound", new Object[]{parts})), 1), 4))) {
            VbaRuntime.unsupportedControlFlow("goto ParseFail");
        }
        if (VbaRuntime.truthy(VbaRuntime.builtin("IsNumeric", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.index(parts, new Object[]{VbaRuntime.builtin("LBound", new Object[]{parts})})})}))) {
            VbaRuntime.set(target, "EntryId", VbaRuntime.builtin("CLng", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.index(parts, new Object[]{VbaRuntime.builtin("LBound", new Object[]{parts})})})}));
        } else {
            VbaRuntime.set(target, "EntryId", fallbackId);
        }
        VbaRuntime.set(target, "AccountCode", VbaRuntime.builtin("UCase$", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.index(parts, new Object[]{VbaRuntime.binary("+", VbaRuntime.builtin("LBound", new Object[]{parts}), 1)})})})}));
        if (VbaRuntime.truthy(VbaRuntime.binary("=", VbaRuntime.builtin("Len", new Object[]{VbaRuntime.get(target, "AccountCode" )}), 0))) {
            VbaRuntime.unsupportedControlFlow("goto ParseFail");
        }
        rawAmt = VbaRuntime.builtin("CDbl", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.index(parts, new Object[]{VbaRuntime.binary("+", VbaRuntime.builtin("LBound", new Object[]{parts}), 2)})})})});
        VbaRuntime.set(target, "Amount", VbaRuntime.builtin("CCur", new Object[]{rawAmt}));
        if (VbaRuntime.truthy(VbaRuntime.builtin("IsDate", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.index(parts, new Object[]{VbaRuntime.binary("+", VbaRuntime.builtin("LBound", new Object[]{parts}), 3)})})})}))) {
            VbaRuntime.set(target, "PostedDate", VbaRuntime.builtin("CDate", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.index(parts, new Object[]{VbaRuntime.binary("+", VbaRuntime.builtin("LBound", new Object[]{parts}), 3)})})})}));
        } else {
            VbaRuntime.set(target, "PostedDate", VbaRuntime.name("Date"));
        }
        VbaRuntime.set(target, "IsReconciled", Boolean.FALSE);
        if (VbaRuntime.truthy(VbaRuntime.binary("and", VbaRuntime.binary(">=", VbaRuntime.binary("+", VbaRuntime.binary("-", VbaRuntime.builtin("UBound", new Object[]{parts}), VbaRuntime.builtin("LBound", new Object[]{parts})), 1), 5), VbaRuntime.binary(">", VbaRuntime.builtin("Len", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.index(parts, new Object[]{VbaRuntime.binary("+", VbaRuntime.builtin("LBound", new Object[]{parts}), 4)})})})}), 0)))) {
            tagTokens = VbaRuntime.builtin("Split", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.index(parts, new Object[]{VbaRuntime.binary("+", VbaRuntime.builtin("LBound", new Object[]{parts}), 4)})}), ","});
            VbaRuntime.unsupportedControlFlow("ReDim target");
            i = VbaRuntime.builtin("LBound", new Object[]{tagTokens});
            for (; VbaRuntime.forCondition(i, VbaRuntime.builtin("UBound", new Object[]{tagTokens}), 1); i = VbaRuntime.binary("+", i, 1)) {
                VbaRuntime.setIndex(VbaRuntime.get(target, "Tags" ), new Object[]{VbaRuntime.binary("+", VbaRuntime.binary("-", i, VbaRuntime.builtin("LBound", new Object[]{tagTokens})), 1)}, VbaRuntime.builtin("LCase$", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.index(tagTokens, new Object[]{i})})})}));
            }
        } else {
            VbaRuntime.unsupportedControlFlow("ReDim target");
            VbaRuntime.setIndex(VbaRuntime.get(target, "Tags" ), new Object[]{1}, "untagged");
        }
        TryParseEntry = Boolean.TRUE;
        if (VbaRuntime.returnNow()) return TryParseEntry;
        // VBA label: ParseFail
        TryParseEntry = Boolean.FALSE;
        return TryParseEntry;
    }

    private void NormalizeEntryInPlace(Object entry) {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        Object i = null;
        Object hasCoreTag = null;
        VbaRuntime.set(entry, "AccountCode", VbaRuntime.builtin("Replace", new Object[]{VbaRuntime.get(entry, "AccountCode" ), " ", "_"}));
        VbaRuntime.set(entry, "AccountCode", VbaRuntime.builtin("Replace", new Object[]{VbaRuntime.get(entry, "AccountCode" ), "-", "_"}));
        if (VbaRuntime.truthy(VbaRuntime.binary("<", VbaRuntime.get(entry, "Amount" ), "0"))) {
            VbaRuntime.set(entry, "IsReconciled", Boolean.FALSE);
        } else if (VbaRuntime.truthy(VbaRuntime.binary(">=", VbaRuntime.get(entry, "Amount" ), "10000"))) {
            VbaRuntime.set(entry, "IsReconciled", Boolean.FALSE);
        } else {
            VbaRuntime.set(entry, "IsReconciled", Boolean.TRUE);
        }
        hasCoreTag = Boolean.FALSE;
        if (VbaRuntime.truthy(VbaRuntime.binary("<>", (VbaRuntime.unary("not", VbaRuntime.unary("not", VbaRuntime.get(entry, "Tags" )))), 0))) {
            i = VbaRuntime.builtin("LBound", new Object[]{VbaRuntime.get(entry, "Tags" )});
            for (; VbaRuntime.forCondition(i, VbaRuntime.builtin("UBound", new Object[]{VbaRuntime.get(entry, "Tags" )}), 1); i = VbaRuntime.binary("+", i, 1)) {
                VbaRuntime.setIndex(VbaRuntime.get(entry, "Tags" ), new Object[]{i}, VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.index(VbaRuntime.get(entry, "Tags" ), new Object[]{i})}));
                if (VbaRuntime.truthy(VbaRuntime.binary("=", VbaRuntime.index(VbaRuntime.get(entry, "Tags" ), new Object[]{i}), "core"))) {
                    hasCoreTag = Boolean.TRUE;
                }
            }
        }
        if (VbaRuntime.truthy(VbaRuntime.unary("not", hasCoreTag))) {
            if (VbaRuntime.truthy(VbaRuntime.binary("=", (VbaRuntime.unary("not", VbaRuntime.unary("not", VbaRuntime.get(entry, "Tags" )))), 0))) {
                VbaRuntime.unsupportedControlFlow("ReDim target");
                VbaRuntime.setIndex(VbaRuntime.get(entry, "Tags" ), new Object[]{1}, "core");
            } else {
                VbaRuntime.unsupportedControlFlow("ReDim target");
                VbaRuntime.setIndex(VbaRuntime.get(entry, "Tags" ), new Object[]{VbaRuntime.builtin("UBound", new Object[]{VbaRuntime.get(entry, "Tags" )})}, "core");
            }
        }
    }

    private void DeduplicateBatchEntries(Object batch) {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        Object seenKeys = null;
        Object kept = null;
        Object keptCount = null;
        Object i = null;
        Object key = null;
        Object newTotal = null;
        VbaRuntime.onError("resume_next", "");
        seenKeys = VbaRuntime.builtin("CreateObject", new Object[]{"Scripting.Dictionary"});
        VbaRuntime.onError("disable_handler", "");
        if (VbaRuntime.truthy(VbaRuntime.binary("is", seenKeys, null))) {
            if (VbaRuntime.returnNow()) return;
        }
        kept = VbaRuntime.redim(kept, new Object[]{VbaRuntime.get(batch, "RecordCount" )});
        keptCount = 0;
        newTotal = "0";
        i = VbaRuntime.builtin("LBound", new Object[]{VbaRuntime.get(batch, "Entries" )});
        for (; VbaRuntime.forCondition(i, VbaRuntime.builtin("UBound", new Object[]{VbaRuntime.get(batch, "Entries" )}), 1); i = VbaRuntime.binary("+", i, 1)) {
            key = VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.get(VbaRuntime.index(VbaRuntime.get(batch, "Entries" ), new Object[]{i}), "EntryId" )}), "::"), VbaRuntime.get(VbaRuntime.index(VbaRuntime.get(batch, "Entries" ), new Object[]{i}), "AccountCode" ));
            if (VbaRuntime.truthy(VbaRuntime.unary("not", VbaRuntime.call(seenKeys, "Exists", new Object[]{key})))) {
                VbaRuntime.call(seenKeys, "Add", new Object[]{key, i});
                keptCount = VbaRuntime.binary("+", keptCount, 1);
                VbaRuntime.setIndex(kept, new Object[]{keptCount}, VbaRuntime.index(VbaRuntime.get(batch, "Entries" ), new Object[]{i}));
                newTotal = VbaRuntime.binary("+", newTotal, VbaRuntime.get(VbaRuntime.index(VbaRuntime.get(batch, "Entries" ), new Object[]{i}), "Amount" ));
            } else {
                this.LogAudit(VbaRuntime.binary("&", "DEDUP: Dropped duplicate key ", key));
            }
        }
        if (VbaRuntime.truthy(VbaRuntime.binary(">", keptCount, 0))) {
            kept = VbaRuntime.redim(kept, new Object[]{keptCount});
            VbaRuntime.set(batch, "Entries", kept);
        } else {
            VbaRuntime.unsupportedControlFlow("Erase target");
        }
        VbaRuntime.set(batch, "RecordCount", keptCount);
        VbaRuntime.set(batch, "TotalAmount", newTotal);
    }

    private void RunLegacyGoSubAudit(Object batch, Object summaryStatus) {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        Object i = null;
        Object flaggedCount = null;
        Object currentCheckId = null;
        Object checkResult = null;
        flaggedCount = 0;
        i = VbaRuntime.builtin("LBound", new Object[]{VbaRuntime.get(batch, "Entries" )});
        for (; VbaRuntime.forCondition(i, VbaRuntime.builtin("UBound", new Object[]{VbaRuntime.get(batch, "Entries" )}), 1); i = VbaRuntime.binary("+", i, 1)) {
            currentCheckId = VbaRuntime.get(VbaRuntime.index(VbaRuntime.get(batch, "Entries" ), new Object[]{i}), "EntryId" );
            VbaRuntime.unsupportedControlFlow("gosub ValidateSingleEntry");
            if (VbaRuntime.truthy(VbaRuntime.binary("<>", checkResult, "OK"))) {
                flaggedCount = VbaRuntime.binary("+", flaggedCount, 1);
                this.LogAudit(VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", "FLAG [", VbaRuntime.builtin("CStr", new Object[]{currentCheckId})), "]: "), checkResult));
            }
        }
        summaryStatus = VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", "AUDITED_", VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.get(batch, "RecordCount" )})), "_FLAGGED_"), VbaRuntime.builtin("CStr", new Object[]{flaggedCount}));
        if (VbaRuntime.returnNow()) return;
        // VBA label: ValidateSingleEntry
        if (VbaRuntime.truthy(VbaRuntime.binary("=", VbaRuntime.get(VbaRuntime.index(VbaRuntime.get(batch, "Entries" ), new Object[]{i}), "Amount" ), "0"))) {
            checkResult = "ZERO_AMOUNT";
        } else if (VbaRuntime.truthy(VbaRuntime.binary("or", VbaRuntime.binary("<", VbaRuntime.builtin("Year", new Object[]{VbaRuntime.get(VbaRuntime.index(VbaRuntime.get(batch, "Entries" ), new Object[]{i}), "PostedDate" )}), 1990), VbaRuntime.binary(">", VbaRuntime.builtin("Year", new Object[]{VbaRuntime.get(VbaRuntime.index(VbaRuntime.get(batch, "Entries" ), new Object[]{i}), "PostedDate" )}), 2099)))) {
            checkResult = "OUT_OF_RANGE_DATE";
        } else {
            checkResult = "OK";
        }
        VbaRuntime.unsupportedControlFlow("gosub_return gosub_return");
    }

    public Object EvaluateDependencyDepth(Object edgesSpec, Object startNode, Object currentDepth, Object maxDepth) {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        Object pairs = null;
        Object edgeParts = null;
        Object i = null;
        Object srcNode = null;
        Object dstNode = null;
        Object childDepth = null;
        Object maxChildDepth = null;
        Object EvaluateDependencyDepth = null;
        if (VbaRuntime.truthy(VbaRuntime.binary(">=", currentDepth, maxDepth))) {
            EvaluateDependencyDepth = currentDepth;
            if (VbaRuntime.returnNow()) return EvaluateDependencyDepth;
        }
        if (VbaRuntime.truthy(VbaRuntime.binary("or", VbaRuntime.binary("=", VbaRuntime.builtin("Len", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{edgesSpec})}), 0), VbaRuntime.binary("=", VbaRuntime.builtin("Len", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{startNode})}), 0)))) {
            EvaluateDependencyDepth = currentDepth;
            if (VbaRuntime.returnNow()) return EvaluateDependencyDepth;
        }
        pairs = VbaRuntime.builtin("Split", new Object[]{edgesSpec, ","});
        maxChildDepth = currentDepth;
        i = VbaRuntime.builtin("LBound", new Object[]{pairs});
        for (; VbaRuntime.forCondition(i, VbaRuntime.builtin("UBound", new Object[]{pairs}), 1); i = VbaRuntime.binary("+", i, 1)) {
            edgeParts = VbaRuntime.builtin("Split", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.index(pairs, new Object[]{i})}), ">"});
            if (VbaRuntime.truthy(VbaRuntime.binary("=", VbaRuntime.binary("+", VbaRuntime.binary("-", VbaRuntime.builtin("UBound", new Object[]{edgeParts}), VbaRuntime.builtin("LBound", new Object[]{edgeParts})), 1), 2))) {
                srcNode = VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.index(edgeParts, new Object[]{VbaRuntime.builtin("LBound", new Object[]{edgeParts})})})});
                dstNode = VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.index(edgeParts, new Object[]{VbaRuntime.binary("+", VbaRuntime.builtin("LBound", new Object[]{edgeParts}), 1)})})});
                if (VbaRuntime.truthy(VbaRuntime.binary("and", VbaRuntime.binary("=", VbaRuntime.builtin("StrComp", new Object[]{srcNode, VbaRuntime.builtin("Trim$", new Object[]{startNode}), VbaRuntime.name("vbTextCompare")}), 0), VbaRuntime.binary(">", VbaRuntime.builtin("Len", new Object[]{dstNode}), 0)))) {
                    if (VbaRuntime.truthy(VbaRuntime.binary("<>", VbaRuntime.builtin("StrComp", new Object[]{dstNode, srcNode, VbaRuntime.name("vbTextCompare")}), 0))) {
                        childDepth = this.EvaluateDependencyDepth(edgesSpec, dstNode, VbaRuntime.binary("+", currentDepth, 1), maxDepth);
                        if (VbaRuntime.truthy(VbaRuntime.binary(">", childDepth, maxChildDepth))) {
                            maxChildDepth = childDepth;
                        }
                    }
                }
            }
        }
        EvaluateDependencyDepth = maxChildDepth;
        return EvaluateDependencyDepth;
    }

    public Object BuildUpsertSql(Object tableName, Object entry) {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        Object cleanTable = null;
        Object cleanAcct = null;
        Object cleanDate = null;
        Object recFlag = null;
        Object BuildUpsertSql = null;
        cleanTable = this.SanitizeSqlIdentifier(tableName);
        cleanAcct = VbaRuntime.builtin("Replace", new Object[]{VbaRuntime.get(entry, "AccountCode" ), "'", "''"});
        cleanDate = VbaRuntime.builtin("Format$", new Object[]{VbaRuntime.get(entry, "PostedDate" ), "yyyy-mm-dd"});
        recFlag = VbaRuntime.builtin("IIf", new Object[]{VbaRuntime.get(entry, "IsReconciled" ), 1, 0});
        BuildUpsertSql = VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", "MERGE INTO ", cleanTable), " AS tgt "), "USING (SELECT "), VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.get(entry, "EntryId" )})), " AS entry_id, '"), cleanAcct), "' AS account_code) AS src "), "ON tgt.entry_id = src.entry_id AND tgt.account_code = src.account_code "), "WHEN MATCHED THEN UPDATE SET amount = "), VbaRuntime.builtin("Replace", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.get(entry, "Amount" )}), ",", "."})), ", posted_date = DATE '"), cleanDate), "', is_reconciled = "), VbaRuntime.builtin("CStr", new Object[]{recFlag})), " "), "WHEN NOT MATCHED THEN INSERT (entry_id, account_code, amount, posted_date, is_reconciled) "), "VALUES ("), VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.get(entry, "EntryId" )})), ", '"), cleanAcct), "', "), VbaRuntime.builtin("Replace", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.get(entry, "Amount" )}), ",", "."})), ", DATE '"), cleanDate), "', "), VbaRuntime.builtin("CStr", new Object[]{recFlag})), ");");
        return BuildUpsertSql;
    }

    private Object SanitizeSqlIdentifier(Object ident) {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        Object i = null;
        Object ch = null;
        Object outStr = null;
        Object SanitizeSqlIdentifier = null;
        outStr = "";
        i = 1;
        for (; VbaRuntime.forCondition(i, VbaRuntime.builtin("Len", new Object[]{ident}), 1); i = VbaRuntime.binary("+", i, 1)) {
            ch = VbaRuntime.builtin("Mid$", new Object[]{ident, i, 1});
            if (VbaRuntime.truthy(VbaRuntime.binary("or", VbaRuntime.binary("or", VbaRuntime.binary("or", VbaRuntime.binary("or", (VbaRuntime.binary("and", VbaRuntime.binary(">=", ch, "A"), VbaRuntime.binary("<=", ch, "Z"))), (VbaRuntime.binary("and", VbaRuntime.binary(">=", ch, "a"), VbaRuntime.binary("<=", ch, "z")))), (VbaRuntime.binary("and", VbaRuntime.binary(">=", ch, "0"), VbaRuntime.binary("<=", ch, "9")))), VbaRuntime.binary("=", ch, "_")), VbaRuntime.binary("=", ch, ".")))) {
                outStr = VbaRuntime.binary("&", outStr, ch);
            }
        }
        if (VbaRuntime.truthy(VbaRuntime.binary("=", VbaRuntime.builtin("Len", new Object[]{outStr}), 0))) {
            outStr = "default_ledger_table";
        }
        SanitizeSqlIdentifier = outStr;
        return SanitizeSqlIdentifier;
    }

    private Object BuildSummaryReport(Object batch, Object parsedCount, Object statusMsg) {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        Object sb = null;
        Object i = null;
        Object logLimit = null;
        Object BuildSummaryReport = null;
        sb = VbaRuntime.binary("&", "=== BATCH SUMMARY ===", VbaRuntime.name("vbCrLf"));
        sb = VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", sb, "Batch ID:      "), VbaRuntime.get(batch, "BatchId" )), VbaRuntime.name("vbCrLf"));
        sb = VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", sb, "Mode Mask:     "), VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.get(batch, "ModeMask" )})), VbaRuntime.name("vbCrLf"));
        sb = VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", sb, "Parsed Rows:   "), VbaRuntime.builtin("CStr", new Object[]{parsedCount})), VbaRuntime.name("vbCrLf"));
        sb = VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", sb, "Final Rows:    "), VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.get(batch, "RecordCount" )})), VbaRuntime.name("vbCrLf"));
        sb = VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", sb, "Total Amount:  "), VbaRuntime.builtin("Format$", new Object[]{VbaRuntime.get(batch, "TotalAmount" ), "#,##0.00"})), VbaRuntime.name("vbCrLf"));
        sb = VbaRuntime.binary("&", VbaRuntime.binary("&", VbaRuntime.binary("&", sb, "Audit Status:  "), statusMsg), VbaRuntime.name("vbCrLf"));
        if (VbaRuntime.truthy(VbaRuntime.binary("and", VbaRuntime.unary("not", VbaRuntime.binary("is", m_AuditLog, null)), VbaRuntime.binary(">", VbaRuntime.get(m_AuditLog, "Count" ), 0)))) {
            sb = VbaRuntime.binary("&", VbaRuntime.binary("&", sb, "--- RECENT AUDIT LOGS ---"), VbaRuntime.name("vbCrLf"));
            logLimit = VbaRuntime.get(m_AuditLog, "Count" );
            if (VbaRuntime.truthy(VbaRuntime.binary(">", logLimit, 15))) {
                logLimit = 15;
            }
            i = 1;
            for (; VbaRuntime.forCondition(i, logLimit, 1); i = VbaRuntime.binary("+", i, 1)) {
                sb = VbaRuntime.binary("&", VbaRuntime.binary("&", sb, VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.index(m_AuditLog, new Object[]{i})})), VbaRuntime.name("vbCrLf"));
            }
        }
        BuildSummaryReport = sb;
        return BuildSummaryReport;
    }

    public Object GetAuditEntryCount() {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        Object GetAuditEntryCount = null;
        if (VbaRuntime.truthy(VbaRuntime.binary("is", m_AuditLog, null))) {
            GetAuditEntryCount = 0;
        } else {
            GetAuditEntryCount = VbaRuntime.get(m_AuditLog, "Count" );
        }
        return GetAuditEntryCount;
    }

    public void ResetModuleState() {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        m_GlobalSequence = 0;
        m_AuditLog = VbaRuntime.newObject("Collection");
        m_NodeRegistry = null;
    }


    private static final class VbaRuntime {
        static final Object MISSING = new Object();
        private static final ErrorState ERR = new ErrorState();

        static final class NamedArg {
            final String name;
            final Object value;
            NamedArg(String name, Object value) { this.name = name; this.value = value; }
        }

        static final class HostObject {
            final String name;
            HostObject(String name) { this.name = name; }
            @Override public String toString() { return name; }
        }

        static final class ErrorState {
            int number;
            String description = "";
            void clear() { number = 0; description = ""; }
            void raise(Object number, Object source, Object description) {
                this.number = (int) VbaRuntime.number(number);
                this.description = String.valueOf(description);
                throw new IllegalStateException(this.description);
            }
        }

        static Object errorState() { return ERR; }
        static Object host(String name) { return new HostObject(name); }
        static Object name(String name) {
            return switch (name.toLowerCase()) {
                case "true" -> Boolean.TRUE;
                case "false" -> Boolean.FALSE;
                case "nothing", "null", "empty" -> null;
                case "vbcrlf", "vbnewline" -> "\r\n";
                case "vbcr" -> "\r";
                case "vblf" -> "\n";
                case "vbtab" -> "\t";
                case "vbnullstring" -> "";
                case "vbtextcompare" -> 1;
                case "vbbinarycompare" -> 0;
                case "vbobjecterror" -> -2147221504;
                case "xlup" -> -4162;
                case "xlyes" -> 1;
                case "xlnone" -> -4142;
                case "xlopenxmlworkbook" -> 51;
                case "vbyellow" -> 6;
                case "vbinformation" -> 64;
                case "vbexclamation" -> 48;
                case "vbcritical" -> 16;
                case "now" -> LocalDateTime.now();
                case "timer" -> LocalDateTime.now().toLocalTime().toSecondOfDay();
                default -> new HostObject(name);
            };
        }

        static Object get(Object target, String member) {
            if (target == null) return null;
            if (target == ERR) {
                return switch (member.toLowerCase()) {
                    case "number" -> ERR.number;
                    case "description" -> ERR.description;
                    default -> null;
                };
            }
            if (target instanceof HostObject host) return new HostObject(host.name + "." + member);
            if (target instanceof Map<?, ?> map) return map.get(member);
            String suffix = Character.toUpperCase(member.charAt(0)) + member.substring(1);
            for (String candidate : new String[]{member, "get" + suffix, "is" + suffix}) {
                try {
                    Method method = target.getClass().getMethod(candidate);
                    return method.invoke(target);
                } catch (ReflectiveOperationException ignored) { }
            }
            try { return target.getClass().getField(member).get(target); }
            catch (ReflectiveOperationException ignored) {
                throw new IllegalStateException("Cannot read VBA member " + member + " from " + target.getClass().getName());
            }
        }

        static Object set(Object target, String member, Object value) {
            if (target == null || target instanceof HostObject)
                throw new UnsupportedOperationException("VBA host adapter required for member " + member);
            if (target instanceof Map<?, ?> raw) {
                @SuppressWarnings("unchecked") Map<Object, Object> map = (Map<Object, Object>) raw;
                map.put(member, value);
                return value;
            }
            String setter = "set" + Character.toUpperCase(member.charAt(0)) + member.substring(1);
            for (Method method : target.getClass().getMethods()) {
                if (method.getName().equalsIgnoreCase(setter) && method.getParameterCount() == 1) {
                    try { method.invoke(target, value); return value; }
                    catch (ReflectiveOperationException exception) { throw new IllegalStateException(exception); }
                }
            }
            throw new IllegalStateException("Cannot write VBA member " + member + " on " + target.getClass().getName());
        }

        static Object call(Object target, String member, Object[] arguments) {
            if (target == ERR && member.equalsIgnoreCase("raise")) {
                Object[] values = unwrap(arguments);
                ERR.raise(
                    values.length > 0 ? values[0] : 0,
                    values.length > 1 ? values[1] : "",
                    values.length > 2 ? values[2] : ""
                );
                return null;
            }
            if (target == ERR && member.equalsIgnoreCase("clear")) {
                ERR.clear();
                return null;
            }
            if (target == null || target instanceof HostObject)
                throw new UnsupportedOperationException("VBA host/API adapter required for " + member);
            Object[] values = unwrap(arguments);
            for (Method method : target.getClass().getMethods()) {
                if (!method.getName().equalsIgnoreCase(member) || method.getParameterCount() != values.length) continue;
                try { return method.invoke(target, values); }
                catch (ReflectiveOperationException exception) { throw new IllegalStateException(exception); }
            }
            throw new IllegalStateException("Cannot invoke VBA member " + member + " on " + target.getClass().getName());
        }

        static Object[] unwrap(Object[] arguments) {
            Object[] values = new Object[arguments.length];
            for (int i = 0; i < arguments.length; i++)
                values[i] = arguments[i] instanceof NamedArg named ? named.value : arguments[i];
            return values;
        }

        static Object builtin(String rawName, Object[] arguments) {
            Object[] args = unwrap(arguments);
            Object first = args.length == 0 ? null : args[0];
            String name = rawName.toLowerCase().replace("$", "");
            return switch (name) {
                case "cstr", "str" -> String.valueOf(first);
                case "len" -> String.valueOf(first == null ? "" : first).length();
                case "trim" -> String.valueOf(first == null ? "" : first).trim();
                case "ltrim" -> String.valueOf(first == null ? "" : first).replaceFirst("^\\s+", "");
                case "rtrim" -> String.valueOf(first == null ? "" : first).replaceFirst("\\s+$", "");
                case "ucase" -> String.valueOf(first == null ? "" : first).toUpperCase();
                case "lcase" -> String.valueOf(first == null ? "" : first).toLowerCase();
                case "isnumeric" -> isNumeric(first);
                case "isdate" -> canParseDate(first);
                case "isarray" -> first != null && (first.getClass().isArray() || first instanceof List<?>);
                case "isnull", "isempty" -> first == null;
                case "cbool" -> truthy(first);
                case "clng", "cint", "cbyte" -> (long) number(first);
                case "cdbl", "csng", "val", "abs" -> name.equals("abs") ? Math.abs(number(first)) : number(first);
                case "ccur", "cdec" -> new BigDecimal(String.valueOf(first));
                case "cdate", "datevalue" -> parseDate(first);
                case "split" -> Arrays.asList(String.valueOf(first).split(
                    args.length > 1 ? java.util.regex.Pattern.quote(String.valueOf(args[1])) : "\\s+", -1));
                case "join" -> String.join(args.length > 1 ? String.valueOf(args[1]) : " ",
                    iterable(args.length == 0 ? null : args[0]).stream().map(String::valueOf).toList());
                case "array" -> args;
                case "ubound" -> upperBound(first);
                case "lbound" -> 0;
                case "replace" -> String.valueOf(first).replace(
                    args.length > 1 ? String.valueOf(args[1]) : "",
                    args.length > 2 ? String.valueOf(args[2]) : "");
                case "left" -> String.valueOf(first).substring(0,
                    Math.min(String.valueOf(first).length(), (int) number(args.length > 1 ? args[1] : 0)));
                case "right" -> {
                    String text = String.valueOf(first);
                    int count = Math.min(text.length(), (int) number(args.length > 1 ? args[1] : 0));
                    yield text.substring(text.length() - count);
                }
                case "mid" -> {
                    String text = String.valueOf(first);
                    int start = Math.max(0, (int) number(args.length > 1 ? args[1] : 1) - 1);
                    int end = args.length > 2 ? Math.min(text.length(), start + (int) number(args[2])) : text.length();
                    yield start >= text.length() ? "" : text.substring(start, end);
                }
                case "iif" -> args.length > 1 && truthy(args[0]) ? args[1] : args.length > 2 ? args[2] : null;
                case "now" -> LocalDateTime.now();
                case "date" -> java.time.LocalDate.now().toString();
                case "year" -> parseDate(first).getYear();
                case "strcomp" -> {
                    String left = String.valueOf(first);
                    String right = args.length > 1 ? String.valueOf(args[1]) : "";
                    boolean textCompare = args.length > 2 && number(args[2]) == 1;
                    yield Integer.signum(textCompare
                        ? left.compareToIgnoreCase(right)
                        : left.compareTo(right));
                }
                case "format" -> format(first, args.length > 1 ? String.valueOf(args[1]) : "");
                case "doevents" -> null;
                case "msgbox" -> { System.out.println(first); yield 0; }
                case "createobject" -> new HostObject("COM:" + String.valueOf(first));
                default -> throw new UnsupportedOperationException("VBA builtin requires implementation: " + rawName);
            };
        }

        static Object binary(String operator, Object left, Object right) {
            return switch (operator.toLowerCase()) {
                case "&" -> String.valueOf(left == null ? "" : left) + String.valueOf(right == null ? "" : right);
                case "+" -> left instanceof String || right instanceof String
                    ? String.valueOf(left) + String.valueOf(right) : number(left) + number(right);
                case "-" -> number(left) - number(right);
                case "*" -> number(left) * number(right);
                case "/" -> number(left) / number(right);
                case "\\" -> (long) (number(left) / number(right));
                case "mod" -> number(left) % number(right);
                case "^" -> Math.pow(number(left), number(right));
                case "=" -> equal(left, right);
                case "<>" -> !equal(left, right);
                case "is" -> left == right;
                case "is_not" -> left != right;
                case "<" -> compare(left, right) < 0;
                case ">" -> compare(left, right) > 0;
                case "<=" -> compare(left, right) <= 0;
                case ">=" -> compare(left, right) >= 0;
                case "and" -> left instanceof Number && right instanceof Number
                    ? ((long) number(left) & (long) number(right)) : truthy(left) && truthy(right);
                case "or" -> left instanceof Number && right instanceof Number
                    ? ((long) number(left) | (long) number(right)) : truthy(left) || truthy(right);
                case "xor" -> left instanceof Number && right instanceof Number
                    ? ((long) number(left) ^ (long) number(right)) : truthy(left) ^ truthy(right);
                case "like" -> String.valueOf(left).matches(
                    java.util.regex.Pattern.quote(String.valueOf(right)).replace("*", "\\E.*\\Q").replace("?", "\\E.\\Q"));
                default -> throw new UnsupportedOperationException("VBA operator requires implementation: " + operator);
            };
        }

        static Object unary(String operator, Object value) {
            return switch (operator.toLowerCase()) {
                case "+" -> number(value);
                case "-" -> -number(value);
                case "not" -> value instanceof Number ? ~((long) number(value)) : !truthy(value);
                default -> throw new UnsupportedOperationException("VBA unary operator requires implementation: " + operator);
            };
        }

        static boolean truthy(Object value) {
            if (value == null || value == MISSING) return false;
            if (value instanceof Boolean bool) return bool;
            if (value instanceof Number number) return number.doubleValue() != 0;
            if (value instanceof String text) return !text.isEmpty();
            return true;
        }

        static boolean equal(Object left, Object right) {
            if (left instanceof Number && right instanceof Number)
                return Double.compare(number(left), number(right)) == 0;
            return Objects.equals(left, right);
        }

        static int compare(Object left, Object right) {
            if (left instanceof Number && right instanceof Number)
                return Double.compare(number(left), number(right));
            return String.valueOf(left).compareTo(String.valueOf(right));
        }

        static double number(Object value) {
            if (value instanceof Number number) return number.doubleValue();
            if (value instanceof Boolean bool) return bool ? -1 : 0;
            if (value == null || String.valueOf(value).isBlank()) return 0;
            try { return Double.parseDouble(String.valueOf(value)); }
            catch (NumberFormatException exception) { throw new IllegalArgumentException("Not numeric: " + value, exception); }
        }

        static boolean isNumeric(Object value) {
            try { number(value); return value != null; }
            catch (RuntimeException exception) { return false; }
        }

        static boolean forCondition(Object current, Object end, Object step) {
            return number(step) >= 0 ? number(current) <= number(end) : number(current) >= number(end);
        }

        static List<Object> iterable(Object value) {
            if (value instanceof List<?> list) return new ArrayList<>(list);
            if (value instanceof Iterable<?> iterable) {
                List<Object> result = new ArrayList<>();
                iterable.forEach(result::add);
                return result;
            }
            if (value != null && value.getClass().isArray()) {
                List<Object> result = new ArrayList<>();
                for (int i = 0; i < Array.getLength(value); i++) result.add(Array.get(value, i));
                return result;
            }
            return Collections.emptyList();
        }

        static Object index(Object value, Object[] indices) {
            int index = indices.length == 0 ? 0 : (int) number(indices[0]);
            if (value instanceof List<?> list) return list.get(index);
            if (value != null && value.getClass().isArray()) return Array.get(value, index);
            if (value instanceof Map<?, ?> map) return map.get(indices.length == 0 ? null : indices[0]);
            throw new IllegalStateException("Cannot index " + value);
        }

        static Object setIndex(Object value, Object[] indices, Object item) {
            int index = indices.length == 0 ? 0 : (int) number(indices[0]);
            if (value instanceof List<?> raw) {
                @SuppressWarnings("unchecked") List<Object> list = (List<Object>) raw;
                while (list.size() <= index) list.add(null);
                list.set(index, item);
                return item;
            }
            if (value != null && value.getClass().isArray()) { Array.set(value, index, item); return item; }
            if (value instanceof Map<?, ?> raw && indices.length > 0) {
                @SuppressWarnings("unchecked") Map<Object, Object> map = (Map<Object, Object>) raw;
                map.put(indices[0], item);
                return item;
            }
            throw new IllegalStateException("Cannot assign indexed value " + value);
        }

        static int upperBound(Object value) {
            if (value instanceof List<?> list) return list.size() - 1;
            if (value != null && value.getClass().isArray()) return Array.getLength(value) - 1;
            return -1;
        }

        static Object redim(Object value, Object[] bounds) {
            int length = bounds.length == 0 ? 0 : Math.max(0, (int) number(bounds[bounds.length - 1]) + 1);
            Object[] result = new Object[length];
            if (value != null && value.getClass().isArray())
                System.arraycopy(value, 0, result, 0, Math.min(Array.getLength(value), result.length));
            return result;
        }

        static Object date(String value) { return value; }
        static LocalDateTime parseDate(Object value) {
            if (value instanceof LocalDateTime dateTime) return dateTime;
            if (value instanceof LocalDate date) return date.atStartOfDay();
            String text = String.valueOf(value);
            for (DateTimeFormatter formatter : new DateTimeFormatter[]{
                    DateTimeFormatter.ISO_LOCAL_DATE_TIME,
                    DateTimeFormatter.ISO_LOCAL_DATE,
                    DateTimeFormatter.ofPattern("M/d/uuuu"),
                    DateTimeFormatter.ofPattern("M/d/uuuu H:mm:ss")}) {
                try {
                    try { return LocalDateTime.parse(text, formatter); }
                    catch (java.time.format.DateTimeParseException ignored) {
                        return LocalDate.parse(text, formatter).atStartOfDay();
                    }
                } catch (java.time.format.DateTimeParseException ignored) { }
            }
            throw new IllegalArgumentException("Unsupported VBA date value: " + value);
        }
        static boolean canParseDate(Object value) {
            if (value == null) return false;
            try { parseDate(value); return true; }
            catch (RuntimeException exception) { return false; }
        }
        static String format(Object value, String vbaPattern) {
            if (!(value instanceof LocalDateTime)) return String.valueOf(value);
            String pattern = vbaPattern.replace("nn", "__MINUTE__")
                .replace("mm", "MM").replace("__MINUTE__", "mm").replace("hh", "HH");
            try { return ((LocalDateTime) value).format(DateTimeFormatter.ofPattern(pattern)); }
            catch (IllegalArgumentException exception) { return String.valueOf(value); }
        }
        static Object newObject(String typeName) { return new HostObject(typeName); }
        static boolean returnNow() { return true; }
        static void onError(String mode, String handler) { }
        static void raiseError(Object code, Object message) {
            throw new IllegalStateException(String.valueOf(message));
        }
        static void unsupportedControlFlow(String description) {
            throw new UnsupportedOperationException("Translate VBA control flow: " + description);
        }
        static Object untranslatedExpression(String source) {
            throw new UnsupportedOperationException("Translate VBA expression: " + source);
        }
        static void untranslated(String source) {
            throw new UnsupportedOperationException("Translate VBA statement: " + source);
        }
    }
}
