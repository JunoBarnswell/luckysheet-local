package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.StructuralStateChanges;
import com.xc.luckysheet.server.contract.GeneratedWorkbookContract;
import com.xc.luckysheet.server.contract.StructuralStateChanges.CellChange;
import com.xc.luckysheet.server.contract.StructuralStateChanges.CellRowChange;
import com.xc.luckysheet.server.contract.StructuralStateChanges.PropertyChange;
import com.xc.luckysheet.server.contract.StructuralStateChanges.PropertyValue;
import com.xc.luckysheet.server.contract.StructuralStateChanges.SheetChange;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

/**
 * Applies complete server-authored facts without re-entering a coordinate or
 * reference reducer. Live planners must record their affected-owner write set
 * directly. Whole-snapshot difference capture is for the explicit historical
 * protocol migration only, not an online planning or commit implementation.
 */
public final class StructuralStateChangeReplay {
    private static final Set<String> WORKBOOK_IDENTITY = Set.of("schema", "version", "unitId", "sheets");

    private StructuralStateChangeReplay() { }

    public static StructuralStateChanges captureForMigration(JsonNode beforeSnapshot, JsonNode afterSnapshot) {
        ObjectNode before = object(beforeSnapshot, "workbook before");
        ObjectNode after = object(afterSnapshot, "workbook after");
        canonicalIdentity(before);
        canonicalIdentity(after);
        for (String field : List.of("schema", "version", "unitId")) {
            if (before.get(field) == null || !before.get(field).equals(after.get(field))) {
                throw invalid("Planner changed workbook identity: " + field);
            }
        }
        Map<String, ObjectNode> beforeSheets = sheets(before);
        Map<String, ObjectNode> afterSheets = sheets(after);
        List<SheetChange> changedSheets = new ArrayList<>();
        for (String sheetId : union(beforeSheets.keySet(), afterSheets.keySet())) {
            ObjectNode left = beforeSheets.get(sheetId);
            ObjectNode right = afterSheets.get(sheetId);
            List<PropertyChange> properties = properties(left, right, Set.of("cells"));
            List<CellRowChange> rows = rows(left == null ? null : object(left.get("cells"), "cells"),
                    right == null ? null : object(right.get("cells"), "cells"));
            if (left == null || right == null || !properties.isEmpty() || !rows.isEmpty()) {
                changedSheets.add(new SheetChange(sheetId, left != null, right != null, properties, rows));
            }
        }
        return new StructuralStateChanges(before.path("unitId").asText(),
                List.copyOf(beforeSheets.keySet()), List.copyOf(afterSheets.keySet()),
                properties(before, after, WORKBOOK_IDENTITY), changedSheets);
    }

    /** Independent result; a failed precondition leaves the caller's snapshot untouched. */
    public static JsonNode apply(JsonNode snapshot, StructuralStateChanges changes) {
        preflight(snapshot, changes);
        ObjectNode result = ((ObjectNode) snapshot).deepCopy();
        write(result, changes);
        return result;
    }

    /** Transaction-owned use only. Every precondition is checked before the first write. */
    public static void applyOnOwnedSnapshot(JsonNode snapshot, StructuralStateChanges changes) {
        preflight(snapshot, changes);
        write((ObjectNode) snapshot, changes);
    }

    public static void preflight(JsonNode snapshot, StructuralStateChanges changes) {
        if (changes == null) throw invalid("Structural facts are required");
        ObjectNode root = object(snapshot, "workbook");
        canonicalIdentity(root);
        if (!changes.unitId().equals(root.path("unitId").asText())) throw conflict("workbook identity");
        Map<String, ObjectNode> current = sheets(root);
        if (!List.copyOf(current.keySet()).equals(changes.sheetOrderBefore())) throw conflict("worksheet order");
        checkProperties(root, changes.workbookProperties(), "workbook");
        for (SheetChange sheet : changes.sheets()) {
            ObjectNode existing = current.get(sheet.sheetId());
            String owner = "sheet " + sheet.sheetId();
            if ((existing != null) != sheet.beforePresent()) throw conflict(owner + " presence");
            checkProperties(existing, sheet.properties(), owner);
            ObjectNode cells = existing == null ? null : object(existing.get("cells"), owner + " cells");
            if (!sheet.afterPresent() && existing != null) {
                if (!fields(existing, Set.of("cells")).equals(propertyNames(sheet.properties()))
                        || !fields(cells, Set.of()).equals(rowNames(sheet.rows()))) {
                    throw conflict(owner + " deletion is missing stored facts");
                }
            }
            for (CellRowChange row : sheet.rows()) {
                JsonNode actualRow = cells == null ? null : cells.get(Integer.toString(row.row()));
                String rowOwner = owner + " row " + row.row();
                if ((actualRow != null) != row.beforePresent()) throw conflict(rowOwner + " presence");
                ObjectNode rowObject = actualRow == null ? null : object(actualRow, rowOwner);
                if (!row.afterPresent() && rowObject != null && rowObject.size() != row.cells().size()) {
                    throw conflict(rowOwner + " deletion is missing stored cells");
                }
                for (CellChange cell : row.cells()) {
                    JsonNode actual = rowObject == null ? null : rowObject.get(Integer.toString(cell.column()));
                    if (!Objects.equals(actual, cell.before())) throw conflict(rowOwner + " column " + cell.column());
                }
            }
        }
    }

    private static void write(ObjectNode root, StructuralStateChanges changes) {
        Map<String, ObjectNode> current = sheets(root);
        writeProperties(root, changes.workbookProperties());
        for (SheetChange change : changes.sheets()) {
            if (!change.afterPresent()) {
                current.remove(change.sheetId());
                continue;
            }
            ObjectNode sheet = current.get(change.sheetId());
            if (sheet == null) {
                sheet = root.objectNode();
                sheet.set("cells", root.objectNode());
                current.put(change.sheetId(), sheet);
            }
            writeProperties(sheet, change.properties());
            ObjectNode cells = (ObjectNode) sheet.get("cells");
            for (CellRowChange changeRow : change.rows()) {
                String rowKey = Integer.toString(changeRow.row());
                if (!changeRow.afterPresent()) {
                    cells.remove(rowKey);
                    continue;
                }
                ObjectNode row = (ObjectNode) cells.get(rowKey);
                if (row == null) {
                    row = cells.objectNode();
                    cells.set(rowKey, row);
                }
                for (CellChange cell : changeRow.cells()) {
                    String column = Integer.toString(cell.column());
                    if (cell.after() == null) row.remove(column);
                    else row.set(column, cell.after().deepCopy());
                }
            }
        }
        ArrayNode ordered = root.arrayNode();
        for (String sheetId : changes.sheetOrderAfter()) ordered.add(current.get(sheetId));
        root.set("sheets", ordered);
    }

    private static void checkProperties(ObjectNode owner, List<PropertyChange> changes, String label) {
        for (PropertyChange change : changes) {
            JsonNode current = owner == null ? null : owner.get(change.name());
            PropertyValue expected = change.before();
            if ((current != null) != expected.present() || (current != null && !current.equals(expected.value()))) {
                throw conflict(label + " property " + change.name());
            }
        }
    }

    private static void writeProperties(ObjectNode owner, List<PropertyChange> changes) {
        for (PropertyChange change : changes) {
            if (change.after().present()) owner.set(change.name(), change.after().value().deepCopy());
            else owner.remove(change.name());
        }
    }

    private static List<PropertyChange> properties(ObjectNode before, ObjectNode after, Set<String> excluded) {
        List<PropertyChange> changes = new ArrayList<>();
        for (String name : union(fields(before, excluded), fields(after, excluded))) {
            JsonNode left = before == null ? null : before.get(name);
            JsonNode right = after == null ? null : after.get(name);
            if (!Objects.equals(left, right)) changes.add(new PropertyChange(name, PropertyValue.of(left), PropertyValue.of(right)));
        }
        return changes;
    }

    private static List<CellRowChange> rows(ObjectNode before, ObjectNode after) {
        List<CellRowChange> changes = new ArrayList<>();
        for (String key : union(fields(before, Set.of()), fields(after, Set.of()))) {
            int row = coordinate(key, 1_048_575);
            JsonNode leftValue = before == null ? null : before.get(key);
            JsonNode rightValue = after == null ? null : after.get(key);
            if (Objects.equals(leftValue, rightValue)) continue;
            ObjectNode left = leftValue == null ? null : object(leftValue, "cell row before");
            ObjectNode right = rightValue == null ? null : object(rightValue, "cell row after");
            List<CellChange> cells = new ArrayList<>();
            for (String columnKey : union(fields(left, Set.of()), fields(right, Set.of()))) {
                int column = coordinate(columnKey, 16_383);
                JsonNode leftCell = left == null ? null : left.get(columnKey);
                JsonNode rightCell = right == null ? null : right.get(columnKey);
                if (!Objects.equals(leftCell, rightCell)) cells.add(new CellChange(column, leftCell, rightCell));
            }
            changes.add(new CellRowChange(row, left != null, right != null, cells));
        }
        return changes;
    }

    private static Map<String, ObjectNode> sheets(ObjectNode root) {
        JsonNode raw = root.get("sheets");
        if (raw == null || !raw.isArray() || raw.isEmpty()) throw invalid("Workbook sheets are missing");
        Map<String, ObjectNode> result = new LinkedHashMap<>();
        for (JsonNode entry : raw) {
            ObjectNode sheet = object(entry, "sheet");
            JsonNode id = sheet.get("id");
            if (id == null || !id.isTextual() || id.textValue().isBlank() || result.put(id.textValue(), sheet) != null) {
                throw invalid("Worksheet identities are missing or repeated");
            }
        }
        return result;
    }

    private static ObjectNode object(JsonNode node, String label) {
        if (node == null || !node.isObject()) throw invalid("Expected object: " + label);
        return (ObjectNode) node;
    }

    private static void canonicalIdentity(ObjectNode snapshot) {
        JsonNode version = snapshot.get("version");
        JsonNode unitId = snapshot.get("unitId");
        if (!GeneratedWorkbookContract.SNAPSHOT_SCHEMA.equals(snapshot.path("schema").asText())
                || version == null || !version.isIntegralNumber() || !version.canConvertToInt()
                || version.intValue() != GeneratedWorkbookContract.SNAPSHOT_VERSION
                || unitId == null || !unitId.isTextual() || unitId.textValue().isBlank()) {
            throw invalid("Structural facts require the canonical workbook snapshot version");
        }
    }

    private static Set<String> fields(ObjectNode node, Set<String> excluded) {
        Set<String> result = new LinkedHashSet<>();
        if (node != null) node.fieldNames().forEachRemaining(name -> { if (!excluded.contains(name)) result.add(name); });
        return result;
    }

    private static Set<String> propertyNames(List<PropertyChange> changes) {
        Set<String> result = new HashSet<>();
        for (PropertyChange change : changes) result.add(change.name());
        return result;
    }

    private static Set<String> rowNames(List<CellRowChange> changes) {
        Set<String> result = new HashSet<>();
        for (CellRowChange change : changes) result.add(Integer.toString(change.row()));
        return result;
    }

    private static <T> Set<T> union(Set<T> left, Set<T> right) {
        Set<T> result = new LinkedHashSet<>(left);
        result.addAll(right);
        return result;
    }

    private static int coordinate(String key, int maximum) {
        try {
            int index = Integer.parseInt(key);
            if (index < 0 || index > maximum || !Integer.toString(index).equals(key)) throw invalid("Non-canonical cell coordinate");
            return index;
        } catch (NumberFormatException error) {
            throw invalid("Non-canonical cell coordinate");
        }
    }

    private static ServiceException invalid(String reason) {
        return new ServiceException("STRUCTURAL_PATCH_INVARIANT", 409, reason);
    }

    private static ServiceException conflict(String owner) {
        return new ServiceException("STRUCTURAL_PATCH_PRECONDITION", 409,
                "Structural fact no longer matches " + owner + "; reload the committed revision and retain the draft");
    }
}
