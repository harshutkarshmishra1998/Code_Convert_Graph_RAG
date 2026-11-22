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

public class vba_sample_5 {
    // Generated from normalized VBA IR. Host/API adapters and VBA compatibility
    // behavior listed in java_codegen_report.json still require implementation.
    private static final Object LOCK_TIMEOUT_SECONDS = 30;
    public void ReconcileInventory(Object inventorySheet, Object movementSheet) {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        Object inventory = null;
        Object movementTotals = null;
        Object lastInventoryRow = null;
        Object lastMovementRow = null;
        Object rowNumber = null;
        Object itemCode = null;
        Object quantity = null;
        Object transactionDate = null;
        Object previousEvents = null;
        Object startedAt = null;
        Object elapsedSeconds = null;
        Object failureNumber = null;
        Object failureDescription = null;
        VbaRuntime.onError("goto_handler", "HandleError");
        previousEvents = VbaRuntime.get(VbaRuntime.host("Application"), "EnableEvents" );
        VbaRuntime.set(VbaRuntime.host("Application"), "EnableEvents", Boolean.FALSE);
        startedAt = VbaRuntime.name("Timer");
        inventory = VbaRuntime.builtin("CreateObject", new Object[]{"Scripting.Dictionary"});
        movementTotals = VbaRuntime.builtin("CreateObject", new Object[]{"Scripting.Dictionary"});
        lastInventoryRow = VbaRuntime.get(VbaRuntime.call(VbaRuntime.call(inventorySheet, "Cells", new Object[]{VbaRuntime.get(VbaRuntime.get(inventorySheet, "Rows" ), "Count" ), "A"}), "End", new Object[]{VbaRuntime.name("xlUp")}), "Row" );
        lastMovementRow = VbaRuntime.get(VbaRuntime.call(VbaRuntime.call(movementSheet, "Cells", new Object[]{VbaRuntime.get(VbaRuntime.get(movementSheet, "Rows" ), "Count" ), "A"}), "End", new Object[]{VbaRuntime.name("xlUp")}), "Row" );
        rowNumber = 2;
        for (; VbaRuntime.forCondition(rowNumber, lastInventoryRow, 1); rowNumber = VbaRuntime.binary("+", rowNumber, 1)) {
            itemCode = VbaRuntime.builtin("UCase$", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.get(VbaRuntime.call(inventorySheet, "Cells", new Object[]{rowNumber, "A"}), "Value2" )})})});
            if (VbaRuntime.truthy(VbaRuntime.binary(">", VbaRuntime.builtin("Len", new Object[]{itemCode}), 0))) {
                if (VbaRuntime.truthy(VbaRuntime.builtin("IsNumeric", new Object[]{VbaRuntime.get(VbaRuntime.call(inventorySheet, "Cells", new Object[]{rowNumber, "B"}), "Value2" )}))) {
                    if (VbaRuntime.truthy(VbaRuntime.call(inventory, "Exists", new Object[]{itemCode}))) {
                        VbaRuntime.call(VbaRuntime.errorState(), "Raise", new Object[]{VbaRuntime.binary("+", VbaRuntime.name("vbObjectError"), 4001), "ReconcileInventory", VbaRuntime.binary("&", "Duplicate inventory item: ", itemCode)});
                    }
                    VbaRuntime.call(inventory, "Add", new Object[]{itemCode, VbaRuntime.builtin("CDbl", new Object[]{VbaRuntime.get(VbaRuntime.call(inventorySheet, "Cells", new Object[]{rowNumber, "B"}), "Value2" )})});
                } else {
                    VbaRuntime.set(VbaRuntime.get(VbaRuntime.call(inventorySheet, "Cells", new Object[]{rowNumber, "B"}), "Interior" ), "Color", VbaRuntime.name("vbRed"));
                }
            }
        }
        rowNumber = 2;
        for (; VbaRuntime.forCondition(rowNumber, lastMovementRow, 1); rowNumber = VbaRuntime.binary("+", rowNumber, 1)) {
            itemCode = VbaRuntime.builtin("UCase$", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.get(VbaRuntime.call(movementSheet, "Cells", new Object[]{rowNumber, "A"}), "Value2" )})})});
            if (VbaRuntime.truthy(VbaRuntime.binary("and", VbaRuntime.builtin("IsDate", new Object[]{VbaRuntime.get(VbaRuntime.call(movementSheet, "Cells", new Object[]{rowNumber, "B"}), "Value2" )}), VbaRuntime.builtin("IsNumeric", new Object[]{VbaRuntime.get(VbaRuntime.call(movementSheet, "Cells", new Object[]{rowNumber, "C"}), "Value2" )})))) {
                transactionDate = VbaRuntime.builtin("CDate", new Object[]{VbaRuntime.get(VbaRuntime.call(movementSheet, "Cells", new Object[]{rowNumber, "B"}), "Value2" )});
                quantity = VbaRuntime.builtin("CDbl", new Object[]{VbaRuntime.get(VbaRuntime.call(movementSheet, "Cells", new Object[]{rowNumber, "C"}), "Value2" )});
                if (VbaRuntime.truthy(VbaRuntime.binary("and", VbaRuntime.binary("<=", VbaRuntime.call(null, "DateDiff", new Object[]{"d", transactionDate, VbaRuntime.name("Date")}), 30), VbaRuntime.binary("<=", transactionDate, VbaRuntime.name("Date"))))) {
                    if (VbaRuntime.truthy(VbaRuntime.call(movementTotals, "Exists", new Object[]{itemCode}))) {
                        VbaRuntime.untranslated("{'kind': 'invocation_expression', 'callee': {'kind': 'identifier', 'name': 'movementTotals', 'binding': {'symbol_kind': 'local', 'symbol_name': 'movementTotals', 'symbol_id': 'node_3ceb96c57a3b274e:local'}, 'inferred_type': {'kind': 'object', 'name': 'Object', 'source_type': 'Object', 'semantics': {'late_bound_possible': True}}}, 'arguments': [{'kind': 'identifier', 'name': 'itemCode', 'binding': {'symbol_kind': 'local', 'symbol_name': 'itemCode', 'symbol_id': 'node_0d6f9d9e422e40b0:local'}, 'inferred_type': {'kind': 'primitive', 'name': 'string', 'source_type': 'String', 'semantics': {'vba_type': 'String'}}}], 'syntax': 'parenthesized', 'source': 'movementTotals(itemCode)', 'invocation_kind': 'unresolved_function_or_array_access'}");
                    } else {
                        VbaRuntime.call(movementTotals, "Add", new Object[]{itemCode, quantity});
                    }
                }
            }
        }
        rowNumber = 2;
        for (; VbaRuntime.forCondition(rowNumber, lastInventoryRow, 1); rowNumber = VbaRuntime.binary("+", rowNumber, 1)) {
            itemCode = VbaRuntime.builtin("UCase$", new Object[]{VbaRuntime.builtin("Trim$", new Object[]{VbaRuntime.builtin("CStr", new Object[]{VbaRuntime.get(VbaRuntime.call(inventorySheet, "Cells", new Object[]{rowNumber, "A"}), "Value2" )})})});
            if (VbaRuntime.truthy(VbaRuntime.call(inventory, "Exists", new Object[]{itemCode}))) {
                quantity = VbaRuntime.call(null, "inventory", new Object[]{itemCode});
                if (VbaRuntime.truthy(VbaRuntime.call(movementTotals, "Exists", new Object[]{itemCode}))) {
                    quantity = VbaRuntime.binary("+", quantity, VbaRuntime.call(null, "movementTotals", new Object[]{itemCode}));
                }
                Object _vba_select_1 = Boolean.TRUE;
                if (VbaRuntime.truthy(VbaRuntime.binary("<", quantity, 0))) {
                    VbaRuntime.set(VbaRuntime.call(inventorySheet, "Cells", new Object[]{rowNumber, "C"}), "Value", "INVESTIGATE");
                } else if (VbaRuntime.truthy(VbaRuntime.binary("<", quantity, 5))) {
                    VbaRuntime.set(VbaRuntime.call(inventorySheet, "Cells", new Object[]{rowNumber, "C"}), "Value", "REORDER");
                } else if (VbaRuntime.truthy(VbaRuntime.binary("=", quantity, 0))) {
                    VbaRuntime.set(VbaRuntime.call(inventorySheet, "Cells", new Object[]{rowNumber, "C"}), "Value", "OUT OF STOCK");
                } else {
                    VbaRuntime.set(VbaRuntime.call(inventorySheet, "Cells", new Object[]{rowNumber, "C"}), "Value", "OK");
                }
                VbaRuntime.set(VbaRuntime.call(inventorySheet, "Cells", new Object[]{rowNumber, "D"}), "Value", quantity);
            }
            elapsedSeconds = VbaRuntime.binary("-", VbaRuntime.name("Timer"), startedAt);
            if (VbaRuntime.truthy(VbaRuntime.binary("<", elapsedSeconds, 0))) {
                elapsedSeconds = VbaRuntime.binary("+", elapsedSeconds, 86400);
            }
            if (VbaRuntime.truthy(VbaRuntime.binary(">", elapsedSeconds, LOCK_TIMEOUT_SECONDS))) {
                VbaRuntime.call(VbaRuntime.errorState(), "Raise", new Object[]{VbaRuntime.binary("+", VbaRuntime.name("vbObjectError"), 4002), "ReconcileInventory", "Reconciliation exceeded its time limit."});
            }
        }
        VbaRuntime.call(VbaRuntime.call(inventorySheet, "Range", new Object[]{VbaRuntime.binary("&", "A1:D", lastInventoryRow)}), "Sort", new Object[]{new VbaRuntime.NamedArg("Key1", VbaRuntime.call(inventorySheet, "Range", new Object[]{"A1"})), new VbaRuntime.NamedArg("Header", VbaRuntime.name("xlYes"))});
        VbaRuntime.unsupportedControlFlow("goto CleanUp");
        // VBA label: HandleError
        failureNumber = VbaRuntime.get(VbaRuntime.errorState(), "Number" );
        failureDescription = VbaRuntime.get(VbaRuntime.errorState(), "Description" );
        this.LogInventoryError(failureNumber, failureDescription, itemCode);
        VbaRuntime.call(null, "MsgBox", new Object[]{VbaRuntime.binary("&", "Inventory reconciliation failed: ", failureDescription), VbaRuntime.name("vbExclamation")});
        // VBA label: CleanUp
        VbaRuntime.set(VbaRuntime.host("Application"), "EnableEvents", previousEvents);
        movementTotals = null;
        inventory = null;
    }

    public void LogInventoryError(Object errorNumber, Object description, Object itemCode) {
        // ByRef/COM/runtime semantics are documented in the semantic report.
        Object logSheet = null;
        Object nextLogRow = null;
        VbaRuntime.onError("resume_next", "");
        logSheet = VbaRuntime.call(VbaRuntime.host("ThisWorkbook"), "Worksheets", new Object[]{"Error Log"});
        if (VbaRuntime.truthy(VbaRuntime.binary("is", logSheet, null))) {
            if (VbaRuntime.returnNow()) return;
        }
        nextLogRow = VbaRuntime.binary("+", VbaRuntime.get(VbaRuntime.call(VbaRuntime.call(logSheet, "Cells", new Object[]{VbaRuntime.get(VbaRuntime.get(logSheet, "Rows" ), "Count" ), "A"}), "End", new Object[]{VbaRuntime.name("xlUp")}), "Row" ), 1);
        VbaRuntime.set(VbaRuntime.call(logSheet, "Cells", new Object[]{nextLogRow, "A"}), "Value", VbaRuntime.name("Now"));
        VbaRuntime.set(VbaRuntime.call(logSheet, "Cells", new Object[]{nextLogRow, "B"}), "Value", errorNumber);
        VbaRuntime.set(VbaRuntime.call(logSheet, "Cells", new Object[]{nextLogRow, "C"}), "Value", description);
        VbaRuntime.set(VbaRuntime.call(logSheet, "Cells", new Object[]{nextLogRow, "D"}), "Value", itemCode);
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
