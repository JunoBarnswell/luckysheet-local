package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Server authority for every persistent operation accepted from a browser.
 *
 * Registration is deliberately stricter than a dispatch map: each accepted
 * mutation carries its parameter reducer, role, server-resolved ranges,
 * locked-range rule, and concurrency rule. Known browser mutations without a
 * proven Java reducer are explicitly unavailable; they are never accepted as
 * opaque JSON and cannot poison replay or checkpoints.
 */
@Component
public class MutationDescriptorRegistry {
    private static final Set<String> CLEAR_MODES = Set.of("all", "contents", "formats", "notes");
    private static final Set<String> KNOWN_MUTATION_IDS = Set.of(
            "automation.recording.changed", "banded.set",
            "cell.restore", "cell.set", "cells.shifted", "cells.shifted.restore",
            "cf.add", "cf.clear", "cf.remove",
            "column.hidden", "column.resize", "column.unhidden", "columns.deleted", "columns.hidden.restore", "columns.inserted", "columns.unhidden.all",
            "comment.add", "comment.remove", "comment.reply", "comment.reply.remove", "comment.resolve",
            "drawing.add", "drawing.anchor", "drawing.payload.update", "drawing.remove", "drawing.transform", "drawing.transform.batch", "drawing.zorder", "drawing.zorder.restore",
            "dv.add", "dv.remove", "filter.remove", "filter.set", "freeze.set",
            "hyperlink.remove", "hyperlink.set", "merge.remove", "merge.set",
            "name.remove", "name.set", "note.remove", "note.set", "note.visibility", "outline.set",
            "pivot.add", "pivot.drilldown.add", "pivot.drilldown.remove", "pivot.refresh", "pivot.remove", "pivot.update",
            "print.area.clear", "print.area.set", "print.document.replace", "print.pageBreak.remove", "print.pageBreak.set", "print.pageBreaks.clear", "print.pageSetup.set",
            "query.definition.replace", "query.load.pivot-source", "query.load.range", "query.load.sheet-table", "query.load.workbook-table",
            "range.clear", "range.clear.restore", "range.paste", "range.set",
            "row.hidden", "row.resize", "row.unhidden", "rows.deleted", "rows.hidden.restore", "rows.inserted", "rows.permuted", "rows.unhidden.all",
            "sheet.add", "sheet.duplicated", "sheet.hidden", "sheet.protect.remove", "sheet.protect.set", "sheet.remove", "sheet.rename", "sheet.reordered", "sheet.restore", "sheet.tabColor", "sheet.unhidden",
            "sheetTable.add", "sheetTable.remove", "sheetTable.update",
            "sparkline.add", "sparkline.group.add", "sparkline.group.remove", "sparkline.group.replace", "sparkline.remove", "sparkline.update",
            "style.set", "table.add", "table.remove", "view.set", "workbook.renamed", "workbook.restore"
    );
    private static final Map<String, String> UNAVAILABLE_REASONS = Map.ofEntries(
            Map.entry("automation.recording.changed", "Recorder state is transient session state and must not enter workbook history."),
            Map.entry("cells.shifted", "Requires one shared reference AST transform and complete structural participant relocation."),
            Map.entry("cells.shifted.restore", "Requires one shared reference AST transform and complete structural participant relocation."),
            Map.entry("rows.inserted", "Requires one shared reference AST transform and complete structural participant relocation."),
            Map.entry("rows.deleted", "Requires one shared reference AST transform and complete structural participant relocation."),
            Map.entry("columns.inserted", "Requires one shared reference AST transform and complete structural participant relocation."),
            Map.entry("columns.deleted", "Requires one shared reference AST transform and complete structural participant relocation."),
            Map.entry("rows.permuted", "Requires row permutation of formulas, objects, tables, validations and every structural participant."),
            Map.entry("sheet.rename", "Requires a shared formula-reference AST rename transform; raw text replacement is forbidden."),
            Map.entry("sheet.duplicated", "Requires identity remapping for scoped names, object payloads, print state and source relationships."),
            Map.entry("sheet.restore", "Client-side live worksheet instances are not a canonical persisted sheet payload."),
            Map.entry("hyperlink.set", "The canonical workbook snapshot currently has no structured hyperlink collection."),
            Map.entry("hyperlink.remove", "The canonical workbook snapshot currently has no structured hyperlink collection."),
            Map.entry("query.load.workbook-table", "Workbook-table query result blocks have no persisted, frontend-readable canonical data plane."),
            Map.entry("workbook.restore", "Only the server restore flow may materialize a historical workbook snapshot.")
    );

    private final Map<String, MutationDescriptor> descriptors = new ConcurrentHashMap<>();

    public MutationDescriptorRegistry() {
        register(new CellDescriptor("cell.set"));
        register(new CellDescriptor("cell.restore"));
        register(new CellDescriptor("range.set"));
        register(new CellDescriptor("range.paste"));
        register(new CellDescriptor("range.clear"));
        register(new CellDescriptor("range.clear.restore"));
        register(new PresentationDescriptor("style.set"));
        register(new PresentationDescriptor("merge.set"));
        register(new PresentationDescriptor("merge.remove"));
        register(new PresentationDescriptor("freeze.set"));
        register(new PresentationDescriptor("row.resize"));
        register(new PresentationDescriptor("column.resize"));
        register(new PresentationDescriptor("view.set"));
        register(new PresentationDescriptor("sheet.hidden"));
        register(new PresentationDescriptor("sheet.unhidden"));
        register(new PresentationDescriptor("sheet.tabColor"));
        register(new ReviewDescriptor("note.set"));
        register(new ReviewDescriptor("note.remove"));
        register(new ReviewDescriptor("note.visibility"));
        register(new ReviewDescriptor("comment.add"));
        register(new ReviewDescriptor("comment.reply"));
        register(new ReviewDescriptor("comment.reply.remove"));
        register(new ReviewDescriptor("comment.resolve"));
        register(new ReviewDescriptor("comment.remove"));
        register(new ProtectionDescriptor("sheet.protect.set"));
        register(new ProtectionDescriptor("sheet.protect.remove"));
        register(new WorkbookRenameDescriptor());
        register(new RestoreDescriptor());
        for (String id : SheetDataMutationDescriptor.IDS) register(new SheetDataMutationDescriptor(id));
        for (String id : DrawingMutationDescriptor.IDS) register(new DrawingMutationDescriptor(id));
        for (String id : PivotMutationDescriptor.IDS) register(new PivotMutationDescriptor(id));
        for (String id : PivotDrillDownMutationDescriptor.IDS) register(new PivotDrillDownMutationDescriptor(id));
        for (String id : SparklineMutationDescriptor.IDS) register(new SparklineMutationDescriptor(id));
        for (String id : WorkbookStateMutationDescriptor.IDS) register(new WorkbookStateMutationDescriptor(id));
        for (String id : QueryMutationDescriptor.IDS) register(new QueryMutationDescriptor(id));
        for (String id : StructuralMutationDescriptor.IDS) register(new StructuralMutationDescriptor(id));
        register(new SheetRenameMutationDescriptor());
        registerUnavailableKnownMutations();
    }

    public void register(MutationDescriptor descriptor) {
        if (descriptors.putIfAbsent(descriptor.id(), descriptor) != null) {
            throw new IllegalStateException("Duplicate mutation descriptor: " + descriptor.id());
        }
    }

    public MutationDescriptor require(String id, boolean internalCall) {
        MutationDescriptor descriptor = descriptors.get(id);
        if (descriptor == null) throw ServiceException.validation("Unknown mutation: " + id);
        if (descriptor.internalOnly() && !internalCall) throw ServiceException.forbidden("Mutation is server-only: " + id);
        if (descriptor instanceof UnavailableDescriptor unavailable) {
            throw ServiceException.unavailable("Mutation is not available for server commit: " + id + ". " + unavailable.reason());
        }
        return descriptor;
    }

    /** Resolve the client-independent mutation policy before reducing it. */
    public MutationPreparation prepare(JsonNode snapshot, OperationMutation mutation, WorkbookAclRole role) {
        MutationDescriptor descriptor = require(mutation.id(), false);
        if (!role.includes(descriptor.requiredRole())) {
            throw ServiceException.forbidden("Workbook role " + descriptor.requiredRole().wireValue() + " is required for mutation " + mutation.id());
        }
        List<RangeRef> ranges = descriptor.affectedRanges(snapshot, mutation);
        if (descriptor.checksProtection()) assertUnlocked(snapshot, ranges, descriptor.protectionAction());
        return new MutationPreparation(descriptor, ranges);
    }

    public JsonNode applyPublicMutations(JsonNode snapshot, List<OperationMutation> mutations) {
        JsonNode current = snapshot.deepCopy();
        for (OperationMutation mutation : mutations) current = require(mutation.id(), false).apply(current, mutation);
        return current;
    }

    public List<RangeRef> resolveRanges(JsonNode snapshot, OperationMutation mutation) {
        return require(mutation.id(), false).affectedRanges(snapshot, mutation);
    }

    public boolean requiresExactBase(Collection<OperationMutation> mutations) {
        return mutations.stream()
                .map(mutation -> require(mutation.id(), false).rebasePolicy())
                .allMatch(MutationRebasePolicy.EXACT_BASE::equals);
    }

    public List<String> ids() {
        return descriptors.keySet().stream().sorted().toList();
    }

    public List<String> acceptedIds() {
        return descriptors.values().stream()
                .filter(descriptor -> !descriptor.internalOnly() && !(descriptor instanceof UnavailableDescriptor))
                .map(MutationDescriptor::id)
                .sorted()
                .toList();
    }

    /** Complete, server-owned explanation for every known mutation not accepted online. */
    public Map<String, String> unavailableReasons() {
        Map<String, String> reasons = new java.util.TreeMap<>();
        for (Map.Entry<String, String> entry : UNAVAILABLE_REASONS.entrySet()) {
            MutationDescriptor descriptor = descriptors.get(entry.getKey());
            if (descriptor instanceof UnavailableDescriptor || (descriptor != null && descriptor.internalOnly())) {
                reasons.put(entry.getKey(), entry.getValue());
            }
        }
        return Map.copyOf(reasons);
    }

    private void registerUnavailableKnownMutations() {
        for (String id : KNOWN_MUTATION_IDS) {
            String reason = UNAVAILABLE_REASONS.get(id);
            if (reason == null && !descriptors.containsKey(id)) {
                throw new IllegalStateException("Known mutation requires an explicit availability classification: " + id);
            }
            if (reason != null) descriptors.putIfAbsent(id, new UnavailableDescriptor(id, reason));
        }
    }

    private void assertUnlocked(JsonNode snapshot, List<RangeRef> ranges, String action) {
        if (ranges.isEmpty()) return;
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        for (JsonNode sheet : SnapshotMutationSupport.sheets(root)) {
            if (!sheet.isObject()) continue;
            JsonNode rules = sheet.path("protectionRules");
            if (!rules.isArray()) continue;
            for (JsonNode rule : rules) {
                if (!rule.isObject() || !rule.path("locked").asBoolean(false) || allows(rule, action)) continue;
                for (RangeRef range : ranges) {
                    if (covers(rule, range)) throw ServiceException.forbidden("Protected area blocks mutation " + action);
                }
            }
        }
    }

    private boolean allows(JsonNode rule, String action) {
        JsonNode actions = rule.path("allowedActions");
        if (!actions.isArray()) return false;
        for (JsonNode item : actions) if (action.equals(item.asText())) return true;
        return false;
    }

    private boolean covers(JsonNode rule, RangeRef range) {
        String scope = rule.path("scope").asText();
        if ("workbook".equals(scope)) return true;
        if ("sheet".equals(scope)) return range.sheetId().equals(rule.path("sheetId").asText());
        if (!"range".equals(scope)) return false;
        JsonNode protectedRange = rule.get("range");
        return protectedRange != null && protectedRange.isObject()
                && range.sheetId().equals(protectedRange.path("sheetId").asText())
                && range.startRow() <= protectedRange.path("endRow").asInt(-1)
                && range.endRow() >= protectedRange.path("startRow").asInt(Integer.MAX_VALUE)
                && range.startColumn() <= protectedRange.path("endColumn").asInt(-1)
                && range.endColumn() >= protectedRange.path("startColumn").asInt(Integer.MAX_VALUE);
    }

    private static abstract class BaseDescriptor implements MutationDescriptor {
        private final String id;
        private final WorkbookAclRole role;
        private final boolean checksProtection;
        private final String action;

        private BaseDescriptor(String id, WorkbookAclRole role, boolean checksProtection, String action) {
            this.id = id;
            this.role = role;
            this.checksProtection = checksProtection;
            this.action = action;
        }

        @Override public String id() { return id; }
        @Override public boolean internalOnly() { return false; }
        @Override public WorkbookAclRole requiredRole() { return role; }
        @Override public MutationRebasePolicy rebasePolicy() { return MutationRebasePolicy.EXACT_BASE; }
        @Override public boolean checksProtection() { return checksProtection; }
        @Override public String protectionAction() { return action; }
    }

    private static final class CellDescriptor extends BaseDescriptor {
        private CellDescriptor(String id) {
            super(id, WorkbookAclRole.EDITOR, true, "edit-cell");
        }

        @Override
        public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
            ObjectNode root = SnapshotMutationSupport.root(snapshot);
            ObjectNode params = SnapshotMutationSupport.params(mutation);
            return switch (id()) {
                case "cell.set", "cell.restore" -> List.of(SnapshotMutationSupport.cellRange(root, mutation.sheetId(), params));
                case "range.clear", "range.clear.restore" -> List.of(requireOwnRange(root, mutation.sheetId(), params));
                case "range.set", "range.paste" -> writeRanges(root, mutation, params);
                default -> throw ServiceException.validation("Unsupported cell mutation: " + id());
            };
        }

        @Override
        public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
            ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
            ObjectNode params = SnapshotMutationSupport.params(mutation);
            ObjectNode sheet = SnapshotMutationSupport.sheet(root, mutation.sheetId());
            switch (id()) {
                case "cell.set" -> setCell(root, sheet, mutation.sheetId(), params);
                case "cell.restore" -> restoreCell(root, sheet, mutation.sheetId(), params);
                case "range.set" -> writeRange(root, sheet, mutation.sheetId(), params, false);
                case "range.paste" -> writeRange(root, sheet, mutation.sheetId(), params, true);
                case "range.clear" -> clearRange(root, sheet, mutation.sheetId(), params);
                case "range.clear.restore" -> restoreRange(root, sheet, mutation.sheetId(), params);
                default -> throw ServiceException.validation("Unsupported cell mutation: " + id());
            }
            return root;
        }

        private List<RangeRef> writeRanges(ObjectNode root, OperationMutation mutation, ObjectNode params) {
            List<RangeRef> ranges = new ArrayList<>();
            RangeRef target = SnapshotMutationSupport.matrixRange(root, mutation.sheetId(), params);
            if (target != null) ranges.add(target);
            if ("range.paste".equals(id()) && params.path("clearSource").asBoolean(false)) ranges.add(SnapshotMutationSupport.range(root, params.get("sourceRange")));
            return List.copyOf(ranges);
        }

        private void setCell(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.coordinate(root, sheetId, params);
            JsonNode value = params.get("value");
            if (value == null || !value.isObject()) throw ServiceException.validation("cell.set value must be an object");
            SnapshotMutationSupport.putCell(sheet, coordinate, value);
        }

        private void restoreCell(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.coordinate(root, sheetId, params);
            JsonNode previous = params.get("previous");
            if (previous == null || previous.isNull()) SnapshotMutationSupport.removeCell(sheet, coordinate);
            else {
                if (!previous.isObject()) throw ServiceException.validation("cell.restore previous must be an object or null");
                SnapshotMutationSupport.putCell(sheet, coordinate, previous);
            }
        }

        private void writeRange(ObjectNode root, ObjectNode targetSheet, String sheetId, ObjectNode params, boolean supportsCut) {
            SnapshotMutationSupport.Matrix matrix = SnapshotMutationSupport.matrix(root, sheetId, params);
            if (supportsCut && params.path("clearSource").asBoolean(false)) {
                RangeRef source = SnapshotMutationSupport.range(root, params.get("sourceRange"));
                SnapshotMutationSupport.clearCells(SnapshotMutationSupport.sheet(root, source.sheetId()), source);
            }
            for (int rowOffset = 0; rowOffset < matrix.values().size(); rowOffset++) {
                ArrayNode row = matrix.values().get(rowOffset);
                for (int columnOffset = 0; columnOffset < row.size(); columnOffset++) {
                    JsonNode value = row.get(columnOffset);
                    if (value == null || value.isNull()) continue;
                    if (!value.isObject()) throw ServiceException.validation("Range values must contain cell objects");
                    SnapshotMutationSupport.putCell(targetSheet, new SnapshotMutationSupport.CellCoordinate(matrix.startRow() + rowOffset, matrix.startColumn() + columnOffset), value);
                }
            }
        }

        private void clearRange(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            RangeRef range = requireOwnRange(root, sheetId, params);
            String mode = params.path("mode").asText("all");
            if ("hyperlinks".equals(mode)) throw ServiceException.unavailable("Hyperlink state is not persisted by the current workbook snapshot");
            if (!CLEAR_MODES.contains(mode)) throw ServiceException.validation("Unsupported clear mode: " + mode);
            if ("notes".equals(mode)) {
                SnapshotMutationSupport.removeNotes(sheet, range);
                return;
            }
            ObjectNode cells = SnapshotMutationSupport.cells(sheet);
            for (int row = range.startRow(); row <= range.endRow(); row++) {
                ObjectNode cellRow = SnapshotMutationSupport.cellRow(cells, row, false);
                if (cellRow == null) continue;
                for (int column = range.startColumn(); column <= range.endColumn(); column++) {
                    String key = Integer.toString(column);
                    JsonNode existing = cellRow.get(key);
                    if (existing == null || !existing.isObject()) continue;
                    if ("all".equals(mode)) cellRow.remove(key);
                    else {
                        ObjectNode next = ((ObjectNode) existing).deepCopy();
                        if ("contents".equals(mode)) {
                            next.putNull("value");
                            next.remove("formula");
                            next.remove("displayValue");
                        } else {
                            next.remove("style");
                            next.remove("styleId");
                            next.remove("numberFormat");
                            next.remove("displayValue");
                        }
                        cellRow.set(key, next);
                    }
                }
                if (cellRow.isEmpty()) cells.remove(Integer.toString(row));
            }
            if ("all".equals(mode)) {
                SnapshotMutationSupport.removeNotes(sheet, range);
                SnapshotMutationSupport.removeThreads(sheet, range);
            }
        }

        private void restoreRange(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            RangeRef range = requireOwnRange(root, sheetId, params);
            SnapshotMutationSupport.clearCells(sheet, range);
            SnapshotMutationSupport.removeNotes(sheet, range);
            SnapshotMutationSupport.removeThreads(sheet, range);
            JsonNode cells = params.get("cells");
            if (cells == null || !cells.isArray()) throw ServiceException.validation("range.clear.restore cells must be an array");
            if (cells.size() > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Range restore is too large");
            for (JsonNode entry : cells) {
                if (!entry.isObject()) throw ServiceException.validation("Range restore cell must be an object");
                SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.coordinate(root, sheetId, (ObjectNode) entry);
                if (!SnapshotMutationSupport.contains(range, coordinate)) throw ServiceException.validation("Range restore cell is outside its range");
                JsonNode value = entry.get("value");
                if (value != null && !value.isNull()) {
                    if (!value.isObject()) throw ServiceException.validation("Range restore value must be an object");
                    SnapshotMutationSupport.putCell(sheet, coordinate, value);
                }
            }
            SnapshotMutationSupport.restoreNotes(root, sheet, sheetId, range, params.get("notes"));
            SnapshotMutationSupport.restoreThreads(root, sheet, sheetId, range, params.get("comments"));
        }

        private RangeRef requireOwnRange(ObjectNode root, String sheetId, ObjectNode params) {
            RangeRef range = SnapshotMutationSupport.range(root, params.get("range"));
            SnapshotMutationSupport.requireSheet(range, sheetId);
            if (SnapshotMutationSupport.cellCount(range) > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Range is too large");
            return range;
        }
    }

    private static final class PresentationDescriptor extends BaseDescriptor {
        private PresentationDescriptor(String id) {
            super(id, WorkbookAclRole.EDITOR, true, "format");
        }

        @Override
        public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
            ObjectNode root = SnapshotMutationSupport.root(snapshot);
            ObjectNode params = SnapshotMutationSupport.params(mutation);
            return switch (id()) {
                case "style.set" -> SnapshotMutationSupport.styleRanges(root, mutation.sheetId(), params);
                case "merge.set", "merge.remove" -> List.of(ownRange(root, mutation.sheetId(), params));
                case "row.resize" -> List.of(SnapshotMutationSupport.rowRange(root, mutation.sheetId(), params));
                case "column.resize" -> List.of(SnapshotMutationSupport.columnRange(root, mutation.sheetId(), params));
                case "freeze.set", "view.set", "sheet.hidden", "sheet.unhidden", "sheet.tabColor" -> List.of();
                default -> throw ServiceException.validation("Unsupported presentation mutation: " + id());
            };
        }

        @Override
        public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
            ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
            ObjectNode params = SnapshotMutationSupport.params(mutation);
            ObjectNode sheet = SnapshotMutationSupport.sheet(root, mutation.sheetId());
            switch (id()) {
                case "style.set" -> style(root, sheet, mutation.sheetId(), params);
                case "merge.set" -> setMerge(root, sheet, mutation.sheetId(), params);
                case "merge.remove" -> removeMerge(root, sheet, mutation.sheetId(), params);
                case "freeze.set" -> freeze(params, sheet);
                case "row.resize" -> resize(root, sheet, mutation.sheetId(), params, "rowHeights", "row", "height");
                case "column.resize" -> resize(root, sheet, mutation.sheetId(), params, "columnWidths", "column", "width");
                case "view.set" -> view(params, sheet);
                case "sheet.hidden" -> sheet.put("hidden", true);
                case "sheet.unhidden" -> sheet.put("hidden", false);
                case "sheet.tabColor" -> tabColor(params, sheet);
                default -> throw ServiceException.validation("Unsupported presentation mutation: " + id());
            }
            return root;
        }

        private void style(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            List<RangeRef> ranges = SnapshotMutationSupport.styleRanges(root, sheetId, params);
            JsonNode style = params.get("style");
            JsonNode numberFormat = params.get("numberFormat");
            if ((style == null || !style.isObject()) && (numberFormat == null || !numberFormat.isTextual())) throw ServiceException.validation("style.set requires style or numberFormat");
            for (RangeRef range : ranges) {
                for (int row = range.startRow(); row <= range.endRow(); row++) {
                    for (int column = range.startColumn(); column <= range.endColumn(); column++) {
                        ObjectNode cell = SnapshotMutationSupport.cell(sheet, new SnapshotMutationSupport.CellCoordinate(row, column), true);
                        if (style != null && style.isObject()) {
                            ObjectNode merged = cell.path("style").isObject() ? ((ObjectNode) cell.get("style")).deepCopy() : cell.objectNode();
                            merged.setAll((ObjectNode) style.deepCopy());
                            cell.set("style", merged);
                        }
                        if (numberFormat != null && numberFormat.isTextual()) cell.set("numberFormat", numberFormat.deepCopy());
                    }
                }
            }
        }

        private void setMerge(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            RangeRef range = ownRange(root, sheetId, params);
            ArrayNode merges = SnapshotMutationSupport.array(sheet, "merges");
            for (int index = merges.size() - 1; index >= 0; index--) if (SnapshotMutationSupport.sameAnchor(merges.get(index).get("range"), range)) merges.remove(index);
            ObjectNode merge = merges.objectNode();
            merge.set("range", SnapshotMutationSupport.rangeNode(range, merges));
            merge.putObject("anchor").put("row", range.startRow()).put("column", range.startColumn());
            merges.add(merge);
        }

        private void removeMerge(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            RangeRef range = ownRange(root, sheetId, params);
            ArrayNode merges = SnapshotMutationSupport.array(sheet, "merges");
            for (int index = 0; index < merges.size(); index++) {
                if (SnapshotMutationSupport.sameAnchor(merges.get(index).get("range"), range)) {
                    merges.remove(index);
                    return;
                }
            }
        }

        private void freeze(ObjectNode params, ObjectNode sheet) {
            JsonNode freeze = params.get("freeze");
            if (freeze == null || !freeze.isObject()) throw ServiceException.validation("freeze.set requires freeze");
            for (String field : List.of("xSplit", "ySplit", "startRow", "startColumn")) {
                if (!freeze.path(field).isIntegralNumber() || freeze.path(field).asInt() < 0) throw ServiceException.validation("freeze." + field + " must be non-negative integer");
            }
            sheet.set("freeze", freeze.deepCopy());
        }

        private void resize(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params, String collection, String coordinateName, String valueName) {
            int coordinate = SnapshotMutationSupport.index(root, sheetId, params, coordinateName);
            JsonNode value = params.get(valueName);
            if (value == null || !value.isNumber() || value.asDouble() <= 0 || !Double.isFinite(value.asDouble())) throw ServiceException.validation(valueName + " must be a positive number");
            SnapshotMutationSupport.object(sheet, collection).set(Integer.toString(coordinate), value.deepCopy());
        }

        private void view(ObjectNode params, ObjectNode sheet) {
            boolean changed = false;
            for (String field : List.of("showGridlines", "showHeaders")) {
                JsonNode value = params.get(field);
                if (value != null) {
                    if (!value.isBoolean()) throw ServiceException.validation(field + " must be boolean");
                    sheet.set(field, value.deepCopy());
                    changed = true;
                }
            }
            JsonNode zoom = params.get("zoom");
            if (zoom != null) {
                if (!zoom.isNumber() || zoom.asDouble() < 25 || zoom.asDouble() > 400) throw ServiceException.validation("zoom must be between 25 and 400");
                sheet.set("zoom", zoom.deepCopy());
                changed = true;
            }
            if (!changed) throw ServiceException.validation("view.set requires a view property");
        }

        private void tabColor(ObjectNode params, ObjectNode sheet) {
            JsonNode color = params.get("color");
            if (color == null || color.isNull()) sheet.remove("tabColor");
            else if (color.isTextual() && color.asText().length() <= 64) sheet.set("tabColor", color.deepCopy());
            else throw ServiceException.validation("tabColor must be a short string or null");
        }

        private RangeRef ownRange(ObjectNode root, String sheetId, ObjectNode params) {
            RangeRef range = SnapshotMutationSupport.range(root, params.get("range"));
            SnapshotMutationSupport.requireSheet(range, sheetId);
            return range;
        }
    }

    private static final class ReviewDescriptor extends BaseDescriptor {
        private ReviewDescriptor(String id) {
            super(id, WorkbookAclRole.COMMENTER, true, "comment");
        }

        @Override
        public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
            ObjectNode root = SnapshotMutationSupport.root(snapshot);
            ObjectNode params = SnapshotMutationSupport.params(mutation);
            return switch (id()) {
                case "note.set", "note.remove", "note.visibility", "comment.add" -> List.of(SnapshotMutationSupport.cellRange(root, mutation.sheetId(), params));
                case "comment.reply", "comment.reply.remove", "comment.resolve", "comment.remove" -> List.of(SnapshotMutationSupport.threadRange(root, mutation.sheetId(), params));
                default -> throw ServiceException.validation("Unsupported review mutation: " + id());
            };
        }

        @Override
        public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
            ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
            ObjectNode params = SnapshotMutationSupport.params(mutation);
            ObjectNode sheet = SnapshotMutationSupport.sheet(root, mutation.sheetId());
            switch (id()) {
                case "note.set" -> setNote(root, sheet, mutation.sheetId(), params);
                case "note.remove" -> removeNote(root, sheet, mutation.sheetId(), params);
                case "note.visibility" -> noteVisibility(root, sheet, mutation.sheetId(), params);
                case "comment.add" -> addThread(root, sheet, mutation.sheetId(), params);
                case "comment.reply" -> addReply(sheet, params);
                case "comment.reply.remove" -> removeReply(sheet, params);
                case "comment.resolve" -> resolveThread(sheet, params);
                case "comment.remove" -> removeThread(sheet, params);
                default -> throw ServiceException.validation("Unsupported review mutation: " + id());
            }
            return root;
        }

        private void setNote(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.coordinate(root, sheetId, params);
            JsonNode note = params.get("note");
            if (note == null || !note.isObject() || note.path("id").asText().isBlank()) throw ServiceException.validation("note.set requires a note with id");
            ArrayNode notes = SnapshotMutationSupport.array(sheet, "notes");
            SnapshotMutationSupport.removeNote(notes, coordinate);
            ObjectNode entry = notes.objectNode();
            entry.put("row", coordinate.row());
            entry.put("column", coordinate.column());
            entry.set("note", note.deepCopy());
            notes.add(entry);
        }

        private void removeNote(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.coordinate(root, sheetId, params);
            if (!SnapshotMutationSupport.removeNote(SnapshotMutationSupport.array(sheet, "notes"), coordinate)) throw ServiceException.notFound("Note not found");
        }

        private void noteVisibility(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.coordinate(root, sheetId, params);
            JsonNode visible = params.get("visible");
            if (visible == null || !visible.isBoolean()) throw ServiceException.validation("note.visibility requires visible");
            ObjectNode entry = SnapshotMutationSupport.findNote(SnapshotMutationSupport.array(sheet, "notes"), coordinate);
            if (entry == null || !entry.path("note").isObject()) throw ServiceException.notFound("Note not found");
            ((ObjectNode) entry.get("note")).set("visible", visible.deepCopy());
        }

        private void addThread(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.coordinate(root, sheetId, params);
            JsonNode thread = params.get("thread");
            if (thread == null || !thread.isObject() || thread.path("id").asText().isBlank()) throw ServiceException.validation("comment.add requires a thread with id");
            if (!sheetId.equals(thread.path("sheetId").asText()) || thread.path("row").asInt(-1) != coordinate.row() || thread.path("column").asInt(-1) != coordinate.column()) {
                throw ServiceException.validation("Comment thread location does not match mutation");
            }
            ArrayNode threads = SnapshotMutationSupport.array(sheet, "commentThreads");
            if (SnapshotMutationSupport.findThread(threads, thread.path("id").asText()) != null) throw ServiceException.conflict("Comment thread already exists");
            threads.add(thread.deepCopy());
        }

        private void addReply(ObjectNode sheet, ObjectNode params) {
            ObjectNode thread = SnapshotMutationSupport.requireThread(sheet, params);
            JsonNode reply = params.get("reply");
            if (reply == null || !reply.isObject() || reply.path("id").asText().isBlank()) throw ServiceException.validation("comment.reply requires a reply with id");
            ArrayNode replies = SnapshotMutationSupport.array(thread, "replies");
            if (SnapshotMutationSupport.findThread(replies, reply.path("id").asText()) != null) throw ServiceException.conflict("Comment reply already exists");
            replies.add(reply.deepCopy());
        }

        private void removeReply(ObjectNode sheet, ObjectNode params) {
            ObjectNode thread = SnapshotMutationSupport.requireThread(sheet, params);
            String replyId = SnapshotMutationSupport.text(params, "replyId");
            ArrayNode replies = SnapshotMutationSupport.array(thread, "replies");
            for (int index = 0; index < replies.size(); index++) {
                if (replyId.equals(replies.get(index).path("id").asText())) {
                    replies.remove(index);
                    return;
                }
            }
            throw ServiceException.notFound("Comment reply not found");
        }

        private void resolveThread(ObjectNode sheet, ObjectNode params) {
            ObjectNode thread = SnapshotMutationSupport.requireThread(sheet, params);
            JsonNode resolved = params.get("resolved");
            if (resolved == null || !resolved.isBoolean()) throw ServiceException.validation("comment.resolve requires resolved");
            if (resolved.asBoolean() && params.path("resolvedAt").asText().isBlank()) throw ServiceException.validation("Resolving a comment requires resolvedAt");
            thread.set("resolved", resolved.deepCopy());
            if (resolved.asBoolean()) thread.set("resolvedAt", params.get("resolvedAt").deepCopy());
            else thread.remove("resolvedAt");
        }

        private void removeThread(ObjectNode sheet, ObjectNode params) {
            String threadId = SnapshotMutationSupport.text(params, "threadId");
            ArrayNode threads = SnapshotMutationSupport.array(sheet, "commentThreads");
            for (int index = 0; index < threads.size(); index++) {
                if (threadId.equals(threads.get(index).path("id").asText())) {
                    threads.remove(index);
                    return;
                }
            }
            throw ServiceException.notFound("Comment thread not found");
        }
    }

    private static final class ProtectionDescriptor extends BaseDescriptor {
        private ProtectionDescriptor(String id) {
            super(id, WorkbookAclRole.OWNER, false, "protect");
        }

        @Override
        public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
            ObjectNode root = SnapshotMutationSupport.root(snapshot);
            ObjectNode params = SnapshotMutationSupport.params(mutation);
            if ("sheet.protect.set".equals(id())) {
                JsonNode rule = params.get("rule");
                if (rule == null || !rule.isObject() || rule.path("id").asText().isBlank()) throw ServiceException.validation("sheet.protect.set requires a rule with id");
                if ("range".equals(rule.path("scope").asText())) return List.of(SnapshotMutationSupport.range(root, rule.get("range")));
            }
            SnapshotMutationSupport.sheet(root, mutation.sheetId());
            return List.of();
        }

        @Override
        public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
            ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
            ObjectNode params = SnapshotMutationSupport.params(mutation);
            ObjectNode sheet = SnapshotMutationSupport.sheet(root, mutation.sheetId());
            ArrayNode rules = SnapshotMutationSupport.array(sheet, "protectionRules");
            if ("sheet.protect.set".equals(id())) {
                JsonNode rule = params.get("rule");
                if (rule == null || !rule.isObject() || rule.path("id").asText().isBlank()) throw ServiceException.validation("sheet.protect.set requires a rule with id");
                String scope = rule.path("scope").asText();
                if (!Set.of("workbook", "sheet", "range").contains(scope) || !rule.path("locked").isBoolean()) throw ServiceException.validation("Protection rule is invalid");
                if ("range".equals(scope)) SnapshotMutationSupport.range(root, rule.get("range"));
                for (int index = 0; index < rules.size(); index++) {
                    if (rule.path("id").asText().equals(rules.get(index).path("id").asText())) {
                        rules.set(index, rule.deepCopy());
                        return root;
                    }
                }
                rules.add(rule.deepCopy());
                return root;
            }
            String ruleId = SnapshotMutationSupport.text(params, "ruleId");
            for (int index = 0; index < rules.size(); index++) {
                if (ruleId.equals(rules.get(index).path("id").asText())) {
                    rules.remove(index);
                    return root;
                }
            }
            throw ServiceException.notFound("Protection rule not found");
        }
    }

    private static final class WorkbookRenameDescriptor extends BaseDescriptor {
        private WorkbookRenameDescriptor() {
            super("workbook.renamed", WorkbookAclRole.EDITOR, false, "");
        }

        @Override public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) { SnapshotMutationSupport.root(snapshot); SnapshotMutationSupport.params(mutation); return List.of(); }

        @Override
        public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
            ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
            String name = SnapshotMutationSupport.text(SnapshotMutationSupport.params(mutation), "name").trim();
            if (name.isBlank() || name.length() > 255) throw ServiceException.validation("Workbook name is invalid");
            root.put("name", name);
            return root;
        }
    }

    private static final class RestoreDescriptor extends BaseDescriptor {
        private RestoreDescriptor() { super("workbook.restore", WorkbookAclRole.OWNER, false, ""); }
        @Override public boolean internalOnly() { return true; }
        @Override public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) { return List.of(); }
        @Override public JsonNode apply(JsonNode snapshot, OperationMutation mutation) { throw ServiceException.forbidden("Workbook restore is generated only by the server restore operation"); }
    }

    private static final class UnavailableDescriptor extends BaseDescriptor {
        private final String reason;

        private UnavailableDescriptor(String id, String reason) {
            super(id, WorkbookAclRole.OWNER, false, "");
            this.reason = reason;
        }

        private String reason() {
            return reason;
        }
        @Override public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) { return List.of(); }
        @Override public JsonNode apply(JsonNode snapshot, OperationMutation mutation) { throw ServiceException.unavailable("Mutation is not available for server commit: " + id() + ". " + reason); }
    }
}
