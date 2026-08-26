package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
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
    private static final Set<String> HORIZONTAL_ALIGNMENTS = Set.of("general", "left", "center", "right", "centerContinuous", "justify", "distributed", "fill");
    private static final Set<String> VERTICAL_ALIGNMENTS = Set.of("top", "middle", "bottom", "justify", "distributed");
    private static final Set<String> READING_ORDERS = Set.of("context", "ltr", "rtl");
    private static final Set<String> TEXT_ORIENTATIONS = Set.of("horizontal", "stacked", "rotateUp", "rotateDown");
    private static final Set<String> CLEAR_FAMILIES = Set.of("all", "contents", "formats", "comments-and-notes", "hyperlinks");
    private static final Set<String> KNOWN_MUTATION_IDS = Set.of(
            "automation.recording.changed", "banded.set",
            "cell.editor.set", "cell.restore", "cell.set", "cells.inserted", "cells.deleted", "cells.inserted.restore", "cells.deleted.restore", "cellTemplate.remove", "cellTemplate.set", "fill.applied", "fill.restored",
            "cf.add", "cf.clear", "cf.remove",
            "column.defaultWidth.resize", "column.hidden", "column.resize", "column.unhidden", "columns.deleted", "columns.hidden.restore", "columns.inserted", "columns.unhidden.all", "columns.visibility",
            "comment.add", "comment.remove", "comment.reply", "comment.reply.remove", "comment.resolve",
            "dataRegion.add", "dataRegion.remove", "dataSource.add", "dataSource.remove", "dataSource.update",
            "drawing.add", "drawing.anchor", "drawing.payload.update", "drawing.remove", "drawing.transform", "drawing.transform.batch", "drawing.zorder", "drawing.zorder.restore",
            "dv.add", "dv.remove", "autoFilter.remove", "autoFilter.set", "freeze.set",
            "hyperlink.remove", "hyperlink.set", "merge.remove", "merge.set",
            "name.remove", "name.set", "note.remove", "note.set", "note.visibility", "outline.set",
            "pivot.add", "pivot.chart.create", "pivot.drilldown.add", "pivot.drilldown.remove", "pivot.refresh", "pivot.remove", "pivot.update",
            "print.area.clear", "print.area.set", "print.document.replace", "print.pageBreak.remove", "print.pageBreak.set", "print.pageBreaks.clear", "print.pageSetup.set",
            "query.definition.replace", "query.load.pivot-source", "query.load.range", "query.load.sheet-table", "query.load.workbook-table",
            "range.clear", "range.clear.restore", "range.paste", "range.set",
            "row.hidden", "row.resize", "row.unhidden", "rows.deleted", "rows.hidden.restore", "rows.inserted", "rows.permuted", "rows.unhidden.all",
            "sheet.add", "sheet.duplicated", "sheet.hidden", "sheet.protect.remove", "sheet.protect.set", "sheet.remove", "sheet.rename", "sheet.reordered", "sheet.restore", "sheet.tabColor", "sheet.unhidden",
            "sheetTable.add", "sheetTable.remove", "sheetTable.update", "sheetTable.autoFilter.set", "tableSheet.update", "ganttSheet.update", "reportSheet.update",
            "sparkline.add", "sparkline.group.add", "sparkline.group.remove", "sparkline.group.replace", "sparkline.remove", "sparkline.update",
            "style.set", "style.preset.set", "format.painter.applied", "cf.reorder", "table.add", "table.remove", "view.set", "workbook.renamed", "workbook.restore",
            "drawing.visibility.set", "drawing.rename"
    );
    private static final Map<String, String> UNAVAILABLE_REASONS = Map.ofEntries(
            Map.entry("automation.recording.changed", "Recorder state is transient session state and must not enter workbook history."),
            Map.entry("rows.inserted", "Requires one shared reference AST transform and complete structural participant relocation."),
            Map.entry("rows.deleted", "Requires one shared reference AST transform and complete structural participant relocation."),
            Map.entry("columns.inserted", "Requires one shared reference AST transform and complete structural participant relocation."),
            Map.entry("columns.deleted", "Requires one shared reference AST transform and complete structural participant relocation."),
            Map.entry("query.load.workbook-table", "Workbook-table query result blocks have no persisted, frontend-readable canonical data plane."),
            Map.entry("pivot.chart.create", "PivotChart is persisted through one canonical drawing.add mutation; the UI command must never cross the workbook mutation boundary."),
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
        register(new PresentationDescriptor("style.preset.set"));
        register(new PresentationDescriptor("format.painter.applied"));
        register(new CellTemplateDescriptor("cellTemplate.set"));
        register(new CellTemplateDescriptor("cellTemplate.remove"));
        for (String id : FillMutationDescriptor.IDS) register(new FillMutationDescriptor(id));
        register(new CellTemplateDescriptor("cell.editor.set"));
        register(new PresentationDescriptor("merge.set"));
        register(new PresentationDescriptor("merge.remove"));
        register(new PresentationDescriptor("freeze.set"));
        register(new PresentationDescriptor("row.resize"));
        register(new PresentationDescriptor("column.resize"));
        register(new PresentationDescriptor("column.defaultWidth.resize"));
        register(new PresentationDescriptor("columns.visibility"));
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
        register(new ConditionalFormatMutationDescriptor("cf.reorder"));
        for (String id : PivotMutationDescriptor.IDS) register(new PivotMutationDescriptor(id));
        for (String id : PivotDrillDownMutationDescriptor.IDS) register(new PivotDrillDownMutationDescriptor(id));
        for (String id : SparklineMutationDescriptor.IDS) register(new SparklineMutationDescriptor(id));
        for (String id : WorkbookStateMutationDescriptor.IDS) register(new WorkbookStateMutationDescriptor(id));
        for (String id : QueryMutationDescriptor.IDS) register(new QueryMutationDescriptor(id));
        for (String id : DataSourceMutationDescriptor.IDS) register(new DataSourceMutationDescriptor(id));
        for (String id : StructuralMutationDescriptor.IDS) register(new StructuralMutationDescriptor(id));
        for (String id : WorkbookStructureMutationDescriptor.IDS) register(new WorkbookStructureMutationDescriptor(id));
        registerUnavailableKnownMutations();
        verifyDeclaredCapabilities();
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
        if (descriptor.checksProtection() && role != WorkbookAclRole.OWNER) {
            ProtectionResolver.assertAllowed(snapshot, ranges, descriptor.protectionAction());
        }
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

    private void verifyDeclaredCapabilities() {
        for (Map.Entry<String, com.xc.luckysheet.server.contract.GeneratedWorkbookContract.MutationCapability> entry
                : com.xc.luckysheet.server.contract.GeneratedWorkbookContract.MUTATIONS.entrySet()) {
            MutationDescriptor descriptor = descriptors.get(entry.getKey());
            if (entry.getValue().remote() && (descriptor == null || descriptor instanceof UnavailableDescriptor || descriptor.internalOnly())) {
                throw new IllegalStateException("Remote contract mutation has no accepted server reducer: " + entry.getKey());
            }
        }
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
                case "range.set" -> writeRanges(root, mutation, params);
                case "range.paste" -> pasteRanges(root, mutation, params);
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
                case "range.paste" -> applyPaste(root, sheet, mutation.sheetId(), params);
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
            if ("range.paste".equals(id()) && params.path("clearSource").asBoolean(false)) ranges.add(requireBoundedSourceRange(root, params));
            return List.copyOf(ranges);
        }

        private List<RangeRef> pasteRanges(ObjectNode root, OperationMutation mutation, ObjectNode params) {
            PasteShape shape = requirePasteShape(root, mutation.sheetId(), params);
            List<RangeRef> ranges = new ArrayList<>();
            ranges.add(shape.target());
            addColumnWidthRanges(root, mutation.sheetId(), SnapshotMutationSupport.requiredObject(params, "snapshot"), shape.allowedRanges(), ranges);
            if (params.path("clearSource").asBoolean(false)) {
                RangeRef source = requireBoundedSourceRange(root, params);
                ranges.add(source);
                if (!mutation.sheetId().equals(source.sheetId())) {
                    addColumnWidthRanges(root, source.sheetId(), SnapshotMutationSupport.requiredObject(params, "sourceSnapshot"), List.of(source), ranges);
                }
            }
            return List.copyOf(ranges);
        }

        private void addColumnWidthRanges(ObjectNode root, String sheetId, ObjectNode snapshot, List<RangeRef> allowedRanges, List<RangeRef> affectedRanges) {
            JsonNode widths = snapshot.get("columnWidths");
            if (widths == null) return;
            if (!widths.isArray()) throw ServiceException.validation("Paste snapshot columnWidths must be an array");
            for (JsonNode entry : widths) {
                if (!entry.isObject()) throw ServiceException.validation("Paste snapshot column width must be an object");
                int column = boundedValue((ObjectNode) entry, "column", SnapshotMutationSupport.MAX_COLUMN);
                if (allowedRanges.stream().noneMatch(range -> column >= range.startColumn() && column <= range.endColumn())) {
                    throw ServiceException.validation("Paste snapshot column width is outside its affected range");
                }
                RangeRef columnRange = new RangeRef(sheetId, 0, SnapshotMutationSupport.MAX_ROW, column, column);
                if (!affectedRanges.contains(columnRange)) affectedRanges.add(columnRange);
            }
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
                RangeRef source = requireBoundedSourceRange(root, params);
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

        private void applyPaste(ObjectNode root, ObjectNode targetSheet, String sheetId, ObjectNode params) {
            PasteShape shape = requirePasteShape(root, sheetId, params);
            applyPasteSnapshot(root, targetSheet, sheetId, SnapshotMutationSupport.requiredObject(params, "snapshot"), shape.allowedRanges());
            if (params.path("clearSource").asBoolean(false)) {
                RangeRef source = requireBoundedSourceRange(root, params);
                if (!sheetId.equals(source.sheetId())) {
                    JsonNode sourceSnapshot = params.get("sourceSnapshot");
                    if (sourceSnapshot == null || !sourceSnapshot.isObject()) throw ServiceException.validation("Cross-sheet paste requires sourceSnapshot");
                    applyPasteSnapshot(root, SnapshotMutationSupport.sheet(root, source.sheetId()), source.sheetId(), (ObjectNode) sourceSnapshot, List.of(source));
                }
            }
        }

        private void applyPasteSnapshot(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode snapshot, List<RangeRef> allowedRanges) {
            JsonNode cells = snapshot.get("cells");
            if (cells == null || !cells.isArray() || cells.size() > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Paste snapshot cells are required and bounded");
            for (JsonNode entry : cells) {
                if (!entry.isObject()) throw ServiceException.validation("Paste snapshot cell must be an object");
                int row = boundedValue((ObjectNode) entry, "row", SnapshotMutationSupport.MAX_ROW);
                int column = boundedValue((ObjectNode) entry, "column", SnapshotMutationSupport.MAX_COLUMN);
                SnapshotMutationSupport.CellCoordinate coordinate = new SnapshotMutationSupport.CellCoordinate(row, column);
                if (allowedRanges.stream().noneMatch(range -> SnapshotMutationSupport.contains(range, coordinate))) throw ServiceException.validation("Paste snapshot cell is outside its affected range");
                JsonNode value = entry.get("value");
                if (value == null || value.isNull()) SnapshotMutationSupport.removeCell(sheet, coordinate);
                else {
                    if (!value.isObject()) throw ServiceException.validation("Paste snapshot cell value must be an object");
                    SnapshotMutationSupport.putCell(sheet, coordinate, value);
                }
            }
            applyPasteNotes(root, sheet, sheetId, snapshot.get("notes"), allowedRanges);
            applyPasteHyperlinks(root, sheet, sheetId, snapshot.get("hyperlinks"), allowedRanges);
            applyPasteComments(root, sheet, sheetId, snapshot.get("commentCells"), snapshot.get("comments"), allowedRanges);
            if (snapshot.has("validations")) {
                if (!snapshot.path("validations").isArray()) throw ServiceException.validation("Paste snapshot validations must be an array");
                assertOwnerSnapshotUnchangedOutsideRanges(root, sheet, sheetId, "dataValidations", snapshot.path("validations"), allowedRanges);
                sheet.set("dataValidations", snapshot.path("validations").deepCopy());
            }
            if (snapshot.has("conditionalFormats")) {
                if (!snapshot.path("conditionalFormats").isArray()) throw ServiceException.validation("Paste snapshot conditionalFormats must be an array");
                assertOwnerSnapshotUnchangedOutsideRanges(root, sheet, sheetId, "conditionalFormats", snapshot.path("conditionalFormats"), allowedRanges);
                sheet.set("conditionalFormats", snapshot.path("conditionalFormats").deepCopy());
            }
            if (snapshot.has("columnWidths")) {
                JsonNode widths = snapshot.get("columnWidths");
                if (!widths.isArray()) throw ServiceException.validation("Paste snapshot columnWidths must be an array");
                ObjectNode targetWidths = SnapshotMutationSupport.object(sheet, "columnWidthsPx");
                for (JsonNode entry : widths) {
                    if (!entry.isObject()) throw ServiceException.validation("Paste snapshot column width must be an object");
                    int column = boundedValue((ObjectNode) entry, "column", SnapshotMutationSupport.MAX_COLUMN);
                    if (allowedRanges.stream().noneMatch(range -> column >= range.startColumn() && column <= range.endColumn())) throw ServiceException.validation("Paste snapshot column width is outside its affected range");
                    String key = Integer.toString(column);
                    JsonNode width = entry.get("widthPx");
                    if (width == null || width.isNull()) targetWidths.remove(key);
                    else {
                        if (!width.isNumber() || width.asDouble() <= 0) throw ServiceException.validation("Paste snapshot width must be positive");
                        targetWidths.set(key, width.deepCopy());
                    }
                }
            }
        }

        private void assertOwnerSnapshotUnchangedOutsideRanges(ObjectNode root, ObjectNode sheet, String sheetId, String property, JsonNode proposed, List<RangeRef> allowedRanges) {
            JsonNode existing = sheet.get(property);
            ArrayNode current = existing != null && existing.isArray() ? (ArrayNode) existing : JsonNodeFactory.instance.arrayNode();
            for (JsonNode rule : current) {
                if (!ownerIsContained(root, sheetId, rule, allowedRanges) && !containsJson(proposed, rule)) throw ServiceException.validation("Paste snapshot changes an unrelated " + property + " rule");
            }
            for (JsonNode rule : proposed) {
                if (!ownerIsContained(root, sheetId, rule, allowedRanges) && !containsJson(current, rule)) throw ServiceException.validation("Paste snapshot adds an unrelated " + property + " rule");
            }
        }

        private boolean ownerIsContained(ObjectNode root, String sheetId, JsonNode rule, List<RangeRef> allowedRanges) {
            if (!rule.isObject()) throw ServiceException.validation("Paste owner rule must be an object");
            JsonNode ranges = rule.get("ranges");
            if (ranges == null || !ranges.isArray() || ranges.isEmpty()) throw ServiceException.validation("Paste owner rule ranges are required");
            for (JsonNode value : ranges) {
                RangeRef range = SnapshotMutationSupport.range(root, value);
                if (!sheetId.equals(range.sheetId())) throw ServiceException.validation("Paste owner rule targets another sheet");
                if (allowedRanges.stream().noneMatch(allowed -> allowed.sheetId().equals(range.sheetId())
                        && allowed.startRow() <= range.startRow() && allowed.endRow() >= range.endRow()
                        && allowed.startColumn() <= range.startColumn() && allowed.endColumn() >= range.endColumn())) return false;
            }
            return true;
        }

        private boolean containsJson(JsonNode array, JsonNode candidate) {
            if (array == null || !array.isArray()) return false;
            for (JsonNode value : array) if (value.equals(candidate)) return true;
            return false;
        }

        private void applyPasteNotes(ObjectNode root, ObjectNode sheet, String sheetId, JsonNode entries, List<RangeRef> allowedRanges) {
            if (entries == null) return;
            if (!entries.isArray()) throw ServiceException.validation("Paste snapshot notes must be an array");
            ArrayNode notes = SnapshotMutationSupport.array(sheet, "notes");
            for (JsonNode entry : entries) {
                if (!entry.isObject()) throw ServiceException.validation("Paste snapshot note must be an object");
                SnapshotMutationSupport.CellCoordinate coordinate = keyCoordinate(root, sheetId, entry.path("key").asText(null));
                if (allowedRanges.stream().noneMatch(range -> SnapshotMutationSupport.contains(range, coordinate))) throw ServiceException.validation("Paste snapshot note is outside its affected range");
                SnapshotMutationSupport.removeNote(notes, coordinate);
                JsonNode value = entry.get("value");
                if (value != null && !value.isNull()) {
                    ObjectNode next = notes.objectNode();
                    next.put("row", coordinate.row());
                    next.put("column", coordinate.column());
                    next.set("note", value.deepCopy());
                    notes.add(next);
                }
            }
        }

        private void applyPasteHyperlinks(ObjectNode root, ObjectNode sheet, String sheetId, JsonNode entries, List<RangeRef> allowedRanges) {
            if (entries == null) return;
            if (!entries.isArray()) throw ServiceException.validation("Paste snapshot hyperlinks must be an array");
            ArrayNode hyperlinks = SnapshotMutationSupport.array(sheet, "hyperlinks");
            for (JsonNode entry : entries) {
                if (!entry.isObject()) throw ServiceException.validation("Paste snapshot hyperlink must be an object");
                SnapshotMutationSupport.CellCoordinate coordinate = keyCoordinate(root, sheetId, entry.path("key").asText(null));
                if (allowedRanges.stream().noneMatch(range -> SnapshotMutationSupport.contains(range, coordinate))) throw ServiceException.validation("Paste snapshot hyperlink is outside its affected range");
                for (int index = hyperlinks.size() - 1; index >= 0; index--) {
                    JsonNode current = hyperlinks.get(index);
                    if (current.path("row").asInt(-1) == coordinate.row() && current.path("column").asInt(-1) == coordinate.column()) hyperlinks.remove(index);
                }
                JsonNode value = entry.get("value");
                if (value != null && !value.isNull()) {
                    ObjectNode next = hyperlinks.objectNode();
                    next.put("row", coordinate.row());
                    next.put("column", coordinate.column());
                    next.set("hyperlink", value.deepCopy());
                    hyperlinks.add(next);
                }
            }
        }

        private void applyPasteComments(ObjectNode root, ObjectNode sheet, String sheetId, JsonNode keys, JsonNode comments, List<RangeRef> allowedRanges) {
            if (keys == null && comments == null) return;
            if (keys != null && !keys.isArray()) throw ServiceException.validation("Paste snapshot commentCells must be an array");
            ArrayNode threads = SnapshotMutationSupport.array(sheet, "commentThreads");
            if (keys != null) for (JsonNode key : keys) {
                SnapshotMutationSupport.CellCoordinate coordinate = keyCoordinate(root, sheetId, key.asText(null));
                if (allowedRanges.stream().noneMatch(range -> SnapshotMutationSupport.contains(range, coordinate))) throw ServiceException.validation("Paste snapshot comment is outside its affected range");
                for (int index = threads.size() - 1; index >= 0; index--) {
                    JsonNode current = threads.get(index);
                    if (current.path("row").asInt(-1) == coordinate.row() && current.path("column").asInt(-1) == coordinate.column()) threads.remove(index);
                }
            }
            if (comments != null) {
                if (!comments.isArray()) throw ServiceException.validation("Paste snapshot comments must be an array");
                for (JsonNode comment : comments) {
                    if (!comment.isObject() || !sheetId.equals(comment.path("sheetId").asText())) throw ServiceException.validation("Paste snapshot comment targets another sheet");
                    int row = boundedValue((ObjectNode) comment, "row", SnapshotMutationSupport.MAX_ROW);
                    int column = boundedValue((ObjectNode) comment, "column", SnapshotMutationSupport.MAX_COLUMN);
                    if (allowedRanges.stream().noneMatch(range -> SnapshotMutationSupport.contains(range, new SnapshotMutationSupport.CellCoordinate(row, column)))) throw ServiceException.validation("Paste snapshot comment is outside its affected range");
                    threads.add(comment.deepCopy());
                }
            }
        }

        private SnapshotMutationSupport.CellCoordinate keyCoordinate(ObjectNode root, String sheetId, String key) {
            if (key == null || !key.matches("\\d+:\\d+")) throw ServiceException.validation("Paste snapshot metadata key is invalid");
            String[] parts = key.split(":");
            ObjectNode params = JsonNodeFactory.instance.objectNode();
            params.put("row", Integer.parseInt(parts[0]));
            params.put("column", Integer.parseInt(parts[1]));
            return SnapshotMutationSupport.coordinate(root, sheetId, params);
        }

        private int boundedValue(ObjectNode object, String property, int maximum) {
            JsonNode value = object.get(property);
            if (value == null || !value.isIntegralNumber() || !value.canConvertToInt()) throw ServiceException.validation(property + " must be an integer");
            int result = value.intValue();
            if (result < 0 || result > maximum) throw ServiceException.validation(property + " is out of bounds");
            return result;
        }

        private PasteShape requirePasteShape(ObjectNode root, String sheetId, ObjectNode params) {
            String transfer = SnapshotMutationSupport.text(params, "transfer");
            if (!Set.of("copy", "move").contains(transfer)) throw ServiceException.validation("Paste transfer is invalid");
            if (!params.path("clearSource").isBoolean()) throw ServiceException.validation("Paste clearSource must be boolean");
            if ("move".equals(transfer) != params.path("clearSource").asBoolean()) throw ServiceException.validation("Paste transfer and clearSource disagree");
            if ("move".equals(transfer) && !params.has("sourceRange")) throw ServiceException.validation("Move paste requires sourceRange");
            if ("copy".equals(transfer) && params.has("sourceRange")) throw ServiceException.validation("Copy paste cannot carry sourceRange");
            ObjectNode spec = SnapshotMutationSupport.requiredObject(params, "spec");
            String content = SnapshotMutationSupport.text(spec, "content");
            String formatting = SnapshotMutationSupport.text(spec, "formatting");
            String operation = SnapshotMutationSupport.text(spec, "operation");
            if (!Set.of("none", "all", "values", "formulas").contains(content)) throw ServiceException.validation("Paste content is invalid");
            if (!Set.of("all", "none", "number-format", "source-formatting", "all-except-borders", "source-theme").contains(formatting)) throw ServiceException.validation("Paste formatting is invalid");
            if (!Set.of("none", "add", "subtract", "multiply", "divide").contains(operation)) throw ServiceException.validation("Paste operation is invalid");
            if ("source-theme".equals(formatting)) throw ServiceException.unavailable("Paste source theme is not supported by the canonical workbook model");
            ObjectNode metadata = SnapshotMutationSupport.requiredObject(spec, "metadata");
            for (String field : List.of("commentsNotes", "validation", "columnWidths", "conditionalFormats", "hyperlinks")) if (!metadata.path(field).isBoolean()) throw ServiceException.validation("Paste metadata field is invalid: " + field);
            if (!spec.path("skipBlanks").isBoolean() || !spec.path("transpose").isBoolean() || !spec.path("link").isBoolean()) throw ServiceException.validation("Paste specification flags are invalid");
            if (spec.path("link").asBoolean(false) && !"none".equals(operation)) throw ServiceException.validation("Paste link cannot combine with arithmetic");
            if ("none".equals(content) && "none".equals(formatting) && "none".equals(operation) && !spec.path("link").asBoolean(false)
                    && !metadata.path("commentsNotes").asBoolean(false) && !metadata.path("validation").asBoolean(false)
                    && !metadata.path("columnWidths").asBoolean(false) && !metadata.path("conditionalFormats").asBoolean(false)
                    && !metadata.path("hyperlinks").asBoolean(false)) throw ServiceException.validation("Paste specification has no effect");
            ObjectNode origin = SnapshotMutationSupport.requiredObject(params, "targetOrigin");
            int row = boundedValue(origin, "row", SnapshotMutationSupport.MAX_ROW);
            int column = boundedValue(origin, "column", SnapshotMutationSupport.MAX_COLUMN);
            ObjectNode extent = SnapshotMutationSupport.requiredObject(params, "sourceExtent");
            int sourceRows = boundedValue(extent, "rows", SnapshotMutationSupport.MAX_ROW + 1);
            int sourceColumns = boundedValue(extent, "columns", SnapshotMutationSupport.MAX_COLUMN + 1);
            if (sourceRows <= 0 || sourceColumns <= 0) throw ServiceException.validation("Paste source extent must be positive");
            boolean transpose = spec.path("transpose").asBoolean(false);
            int rows = transpose ? sourceColumns : sourceRows;
            int columns = transpose ? sourceRows : sourceColumns;
            ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
            int rowCount = sheet.path("rowCount").asInt(SnapshotMutationSupport.MAX_ROW + 1);
            int columnCount = sheet.path("columnCount").asInt(SnapshotMutationSupport.MAX_COLUMN + 1);
            if (row + rows > rowCount || column + columns > columnCount) throw ServiceException.validation("Paste exceeds worksheet bounds");
            return new PasteShape(new RangeRef(sheetId, row, row + rows - 1, column, column + columns - 1), params.path("clearSource").asBoolean(false) ? requireBoundedSourceRange(root, params) : null);
        }

        private record PasteShape(RangeRef target, RangeRef source) {
            private List<RangeRef> allowedRanges() {
                return source == null ? List.of(target) : source.sheetId().equals(target.sheetId()) ? List.of(target, source) : List.of(target);
            }
        }

        private void clearRange(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            RangeRef range = requireOwnRange(root, sheetId, params);
            String family = params.path("family").asText(null);
            if (!CLEAR_FAMILIES.contains(family)) throw ServiceException.validation("Unsupported clear family: " + family);
            if ("comments-and-notes".equals(family)) {
                SnapshotMutationSupport.removeNotes(sheet, range);
                SnapshotMutationSupport.removeThreads(sheet, range);
                return;
            }
            if ("hyperlinks".equals(family)) {
                SnapshotMutationSupport.removeHyperlinks(sheet, range);
                return;
            }
            ObjectNode cells = SnapshotMutationSupport.cells(sheet);
            for (int row = range.startRow(); row <= range.endRow(); row++) {
                ObjectNode cellRow = SnapshotMutationSupport.cellRow(cells, row, false);
                if (cellRow == null) continue;
                List<String> keys = new ArrayList<>();
                cellRow.fieldNames().forEachRemaining(key -> {
                    try {
                        int column = Integer.parseInt(key);
                        if (column >= range.startColumn() && column <= range.endColumn()) keys.add(key);
                    } catch (NumberFormatException ignored) {
                        throw ServiceException.validation("Cell column key is invalid");
                    }
                });
                for (String key : keys) {
                    JsonNode existing = cellRow.get(key);
                    if (existing == null || !existing.isObject()) continue;
                    if ("all".equals(family)) cellRow.remove(key);
                    else {
                        ObjectNode next = ((ObjectNode) existing).deepCopy();
                        if ("contents".equals(family)) {
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
            if ("all".equals(family)) {
                SnapshotMutationSupport.removeNotes(sheet, range);
                SnapshotMutationSupport.removeThreads(sheet, range);
                SnapshotMutationSupport.removeHyperlinks(sheet, range);
            }
            if ("formats".equals(family) || "all".equals(family)) cropConditionalFormats(root, sheet, sheetId, range);
        }

        private void restoreRange(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            RangeRef range = requireOwnRange(root, sheetId, params);
            SnapshotMutationSupport.clearCells(sheet, range);
            SnapshotMutationSupport.removeNotes(sheet, range);
            SnapshotMutationSupport.removeThreads(sheet, range);
            ObjectNode snapshot = SnapshotMutationSupport.requiredObject(params, "snapshot");
            JsonNode cells = snapshot.get("cells");
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
            SnapshotMutationSupport.restoreNotes(root, sheet, sheetId, range, snapshot.get("notes"));
            SnapshotMutationSupport.restoreThreads(root, sheet, sheetId, range, snapshot.get("comments"));
            restoreHyperlinks(root, sheet, sheetId, range, snapshot.get("hyperlinks"));
            if (snapshot.has("conditionalFormats")) {
                JsonNode rules = snapshot.get("conditionalFormats");
                if (!rules.isArray()) throw ServiceException.validation("Range restore conditionalFormats must be an array");
                sheet.set("conditionalFormats", rules.deepCopy());
            }
        }

        private void restoreHyperlinks(ObjectNode root, ObjectNode sheet, String sheetId, RangeRef range, JsonNode value) {
            if (value == null || value.isNull()) return;
            if (!value.isArray()) throw ServiceException.validation("Range restore hyperlinks must be an array");
            ArrayNode hyperlinks = SnapshotMutationSupport.array(sheet, "hyperlinks");
            SnapshotMutationSupport.removeHyperlinks(sheet, range);
            for (JsonNode entry : value) {
                if (!entry.isObject()) throw ServiceException.validation("Range restore hyperlink must be an object");
                ObjectNode object = (ObjectNode) entry;
                SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.coordinate(root, sheetId, object);
                if (!SnapshotMutationSupport.contains(range, coordinate)) throw ServiceException.validation("Range restore hyperlink is outside its range");
                JsonNode hyperlink = object.get("hyperlink");
                if (hyperlink == null || !hyperlink.isObject()) throw ServiceException.validation("Range restore hyperlink payload is invalid");
                ObjectNode next = hyperlinks.objectNode();
                next.put("row", coordinate.row());
                next.put("column", coordinate.column());
                next.set("hyperlink", hyperlink.deepCopy());
                hyperlinks.add(next);
            }
        }

        private void cropConditionalFormats(ObjectNode root, ObjectNode sheet, String sheetId, RangeRef clear) {
            JsonNode existing = sheet.get("conditionalFormats");
            if (existing == null || existing.isNull()) return;
            if (!existing.isArray()) throw ServiceException.validation("conditionalFormats must be an array");
            ArrayNode nextRules = JsonNodeFactory.instance.arrayNode();
            for (JsonNode rule : existing) {
                if (!rule.isObject()) throw ServiceException.validation("conditionalFormats rule must be an object");
                ObjectNode copy = ((ObjectNode) rule).deepCopy();
                JsonNode ranges = copy.get("ranges");
                if (ranges == null || !ranges.isArray()) throw ServiceException.validation("conditionalFormats rule ranges must be an array");
                ArrayNode remaining = JsonNodeFactory.instance.arrayNode();
                for (JsonNode candidate : ranges) {
                    RangeRef source = SnapshotMutationSupport.range(root, candidate);
                    if (!sheetId.equals(source.sheetId())) throw ServiceException.validation("conditionalFormats rule targets another sheet");
                    for (RangeRef part : subtractRange(source, clear)) remaining.add(rangeNode(part));
                }
                if (!remaining.isEmpty()) {
                    copy.set("ranges", remaining);
                    nextRules.add(copy);
                }
            }
            sheet.set("conditionalFormats", nextRules);
        }

        private List<RangeRef> subtractRange(RangeRef source, RangeRef clear) {
            if (source.startRow() > clear.endRow() || clear.startRow() > source.endRow()
                    || source.startColumn() > clear.endColumn() || clear.startColumn() > source.endColumn()) return List.of(source);
            int top = Math.max(source.startRow(), clear.startRow());
            int bottom = Math.min(source.endRow(), clear.endRow());
            int left = Math.max(source.startColumn(), clear.startColumn());
            int right = Math.min(source.endColumn(), clear.endColumn());
            List<RangeRef> result = new ArrayList<>();
            if (source.startRow() < top) result.add(new RangeRef(source.sheetId(), source.startRow(), top - 1, source.startColumn(), source.endColumn()));
            if (bottom < source.endRow()) result.add(new RangeRef(source.sheetId(), bottom + 1, source.endRow(), source.startColumn(), source.endColumn()));
            if (source.startColumn() < left) result.add(new RangeRef(source.sheetId(), top, bottom, source.startColumn(), left - 1));
            if (right < source.endColumn()) result.add(new RangeRef(source.sheetId(), top, bottom, right + 1, source.endColumn()));
            return result;
        }

        private ObjectNode rangeNode(RangeRef range) {
            ObjectNode node = JsonNodeFactory.instance.objectNode();
            node.put("sheetId", range.sheetId());
            node.put("startRow", range.startRow());
            node.put("endRow", range.endRow());
            node.put("startColumn", range.startColumn());
            node.put("endColumn", range.endColumn());
            return node;
        }

        private RangeRef requireOwnRange(ObjectNode root, String sheetId, ObjectNode params) {
            RangeRef range = SnapshotMutationSupport.range(root, params.get("range"));
            SnapshotMutationSupport.requireSheet(range, sheetId);
            if (SnapshotMutationSupport.cellCount(range) > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Range is too large");
            return range;
        }

        private RangeRef requireBoundedSourceRange(ObjectNode root, ObjectNode params) {
            RangeRef range = SnapshotMutationSupport.range(root, params.get("sourceRange"));
            if (SnapshotMutationSupport.cellCount(range) > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Paste source range is too large");
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
                case "style.set", "style.preset.set" -> SnapshotMutationSupport.styleRanges(root, mutation.sheetId(), params);
                case "format.painter.applied" -> List.of(SnapshotMutationSupport.range(root, params.get("targetRange")));
                case "merge.set", "merge.remove" -> List.of(ownRange(root, mutation.sheetId(), params));
                case "row.resize" -> List.of(SnapshotMutationSupport.rowRange(root, mutation.sheetId(), params));
                case "column.resize" -> List.of(SnapshotMutationSupport.columnRange(root, mutation.sheetId(), params));
                case "columns.visibility" -> visibilityRanges(root, mutation.sheetId(), params);
                case "freeze.set", "view.set", "column.defaultWidth.resize", "sheet.hidden", "sheet.unhidden", "sheet.tabColor" -> List.of();
                default -> throw ServiceException.validation("Unsupported presentation mutation: " + id());
            };
        }

        @Override
        public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
            ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
            ObjectNode params = SnapshotMutationSupport.params(mutation);
            ObjectNode sheet = SnapshotMutationSupport.sheet(root, mutation.sheetId());
            switch (id()) {
                case "style.set", "style.preset.set" -> style(root, sheet, mutation.sheetId(), params);
                case "format.painter.applied" -> formatPainter(root, sheet, mutation.sheetId(), params);
                case "merge.set" -> setMerge(root, sheet, mutation.sheetId(), params);
                case "merge.remove" -> removeMerge(root, sheet, mutation.sheetId(), params);
                case "freeze.set" -> freeze(params, sheet);
                case "row.resize" -> resize(root, sheet, mutation.sheetId(), params, "rowHeightsPx", "row", "heightPx");
                case "column.resize" -> resize(root, sheet, mutation.sheetId(), params, "columnWidthsPx", "column", "widthPx");
                case "column.defaultWidth.resize" -> defaultColumnWidth(params, sheet);
                case "columns.visibility" -> columnVisibility(root, sheet, mutation.sheetId(), params);
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
            if (style != null && style.isObject()) validateStyle((ObjectNode) style);
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

        private void formatPainter(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            RangeRef target = SnapshotMutationSupport.range(root, params.get("targetRange"));
            SnapshotMutationSupport.requireSheet(target, sheetId);
            JsonNode styles = params.get("styles");
            if (styles == null || !styles.isArray()) throw ServiceException.validation("format.painter.applied requires styles");
            int height = target.endRow() - target.startRow() + 1;
            int width = target.endColumn() - target.startColumn() + 1;
            if (styles.size() != height) throw ServiceException.validation("format painter row count does not match target range");
            for (int rowOffset = 0; rowOffset < height; rowOffset++) {
                JsonNode row = styles.get(rowOffset);
                if (!row.isArray() || row.size() != width) throw ServiceException.validation("format painter column count does not match target range");
                for (int columnOffset = 0; columnOffset < width; columnOffset++) {
                    JsonNode styleEntry = row.get(columnOffset);
                    if (!styleEntry.isObject()) throw ServiceException.validation("format painter style entry must be an object");
                    ObjectNode cell = SnapshotMutationSupport.cell(sheet, new SnapshotMutationSupport.CellCoordinate(target.startRow() + rowOffset, target.startColumn() + columnOffset), true);
                    JsonNode style = styleEntry.get("style");
                    if (style != null && !style.isNull()) {
                        if (!style.isObject()) throw ServiceException.validation("format painter style must be an object");
                        validateStyle((ObjectNode) style);
                        ObjectNode merged = cell.path("style").isObject() ? ((ObjectNode) cell.get("style")).deepCopy() : cell.objectNode();
                        merged.setAll((ObjectNode) style.deepCopy());
                        cell.set("style", merged);
                    }
                    JsonNode numberFormat = styleEntry.get("numberFormat");
                    if (numberFormat != null && !numberFormat.isNull()) {
                        if (!numberFormat.isTextual()) throw ServiceException.validation("format painter numberFormat must be text");
                        cell.set("numberFormat", numberFormat.deepCopy());
                    }
                }
            }
        }

        private void validateStyle(ObjectNode style) {
            if (style.has("unsupportedAlignment")) throw ServiceException.validation("style.set cannot edit unsupported alignment attributes");
            if (style.has("horizontalAlignment") && !HORIZONTAL_ALIGNMENTS.contains(SnapshotMutationSupport.text(style, "horizontalAlignment"))) throw ServiceException.validation("horizontalAlignment is invalid");
            if (style.has("verticalAlignment") && !VERTICAL_ALIGNMENTS.contains(SnapshotMutationSupport.text(style, "verticalAlignment"))) throw ServiceException.validation("verticalAlignment is invalid");
            if (style.has("readingOrder") && !READING_ORDERS.contains(SnapshotMutationSupport.text(style, "readingOrder"))) throw ServiceException.validation("readingOrder is invalid");
            if (style.has("textOrientation") && !TEXT_ORIENTATIONS.contains(SnapshotMutationSupport.text(style, "textOrientation"))) throw ServiceException.validation("textOrientation is invalid");
            if (style.has("shrinkToFit") && !style.path("shrinkToFit").isBoolean()) throw ServiceException.validation("shrinkToFit is invalid");
            if (style.has("indent") && (!style.path("indent").canConvertToInt() || style.path("indent").asInt() < 0 || style.path("indent").asInt() > 250)) throw ServiceException.validation("indent is invalid");
            if (style.has("textRotate") && (!style.path("textRotate").isNumber() || !Double.isFinite(style.path("textRotate").asDouble()) || style.path("textRotate").asDouble() < -180 || style.path("textRotate").asDouble() > 180)) throw ServiceException.validation("textRotate is invalid");
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
            JsonNode pane = params.get("pane");
            if (pane == null || !pane.isObject()) throw ServiceException.validation("freeze.set requires pane");
            String kind = pane.path("kind").asText();
            if (!Set.of("none", "frozen", "split").contains(kind)) throw ServiceException.validation("pane.kind is invalid");
            if ("none".equals(kind)) { sheet.set("pane", pane.deepCopy()); return; }
            String state = pane.path("state").asText();
            if (("frozen".equals(kind) && !Set.of("frozen", "frozenSplit").contains(state))
                    || ("split".equals(kind) && !"split".equals(state))) {
                throw ServiceException.validation("pane.state is invalid for pane.kind");
            }
            for (String field : List.of("xSplit", "ySplit", "startRow", "startColumn")) {
                if (!pane.path(field).isNumber() || pane.path(field).asDouble() < 0) throw ServiceException.validation("pane." + field + " must be non-negative");
            }
            if (!pane.path("startRow").isIntegralNumber() || !pane.path("startColumn").isIntegralNumber()) throw ServiceException.validation("pane start coordinates must be integers");
            sheet.set("pane", pane.deepCopy());
        }

        private List<RangeRef> visibilityRanges(ObjectNode root, String sheetId, ObjectNode params) {
            JsonNode states = params.get("states");
            if (states == null || !states.isArray() || states.isEmpty()) throw ServiceException.validation("columns.visibility requires states");
            List<RangeRef> ranges = new ArrayList<>();
            for (JsonNode state : states) {
                ObjectNode coordinate = params.objectNode().put("column", state.path("column").asInt(-1));
                ranges.add(SnapshotMutationSupport.columnRange(root, sheetId, coordinate));
            }
            return List.copyOf(ranges);
        }

        private void columnVisibility(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
            ArrayNode states = (ArrayNode) params.path("states");
            ArrayNode hidden = SnapshotMutationSupport.array(sheet, "hiddenColumns");
            java.util.Set<Integer> values = new java.util.TreeSet<>();
            hidden.forEach(value -> { if (value.isIntegralNumber()) values.add(value.asInt()); });
            for (JsonNode state : states) {
                int column = SnapshotMutationSupport.index(root, sheetId, (ObjectNode) state, "column");
                if (!state.path("hidden").isBoolean()) throw ServiceException.validation("columns.visibility hidden must be boolean");
                if (state.path("hidden").asBoolean()) values.add(column); else values.remove(column);
            }
            hidden.removeAll(); values.forEach(hidden::add);
        }

        private void defaultColumnWidth(ObjectNode params, ObjectNode sheet) {
            JsonNode width = params.get("widthPx");
            if (width == null || !width.isNumber() || width.asDouble() <= 0 || !Double.isFinite(width.asDouble())) throw ServiceException.validation("widthPx must be a positive number");
            sheet.set("defaultColumnWidthPx", width.deepCopy());
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

    /** Workbook-native cell template and editor state. Template library writes
     * are workbook scoped; editor writes are range-scoped presentation changes. */
    private static final class CellTemplateDescriptor extends BaseDescriptor {
        private CellTemplateDescriptor(String id) {
            super(id, WorkbookAclRole.EDITOR, "cell.editor.set".equals(id), "format");
        }

        @Override
        public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
            ObjectNode root = SnapshotMutationSupport.root(snapshot);
            ObjectNode params = SnapshotMutationSupport.params(mutation);
            return switch (id()) {
                case "cell.editor.set" -> SnapshotMutationSupport.styleRanges(root, mutation.sheetId(), params);
                case "cellTemplate.set", "cellTemplate.remove" -> List.of();
                default -> throw ServiceException.validation("Unsupported cell template mutation: " + id());
            };
        }

        @Override
        public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
            ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
            ObjectNode params = SnapshotMutationSupport.params(mutation);
            switch (id()) {
                case "cellTemplate.set" -> setTemplate(root, params);
                case "cellTemplate.remove" -> removeTemplate(root, params);
                case "cell.editor.set" -> setEditor(root, mutation.sheetId(), params);
                default -> throw ServiceException.validation("Unsupported cell template mutation: " + id());
            }
            return root;
        }

        private void setTemplate(ObjectNode root, ObjectNode params) {
            JsonNode raw = params.get("template");
            if (raw == null || !raw.isObject()) throw ServiceException.validation("cellTemplate.set requires template");
            String id = raw.path("id").asText();
            String name = raw.path("name").asText();
            if (id.isBlank() || name.isBlank() || !raw.path("style").isObject()) throw ServiceException.validation("Cell template is invalid");
            ArrayNode templates = SnapshotMutationSupport.array(root, "cellStyleTemplates");
            for (int index = templates.size() - 1; index >= 0; index--) {
                if (id.equals(templates.get(index).path("id").asText())) templates.remove(index);
            }
            templates.add(raw.deepCopy());
        }

        private void removeTemplate(ObjectNode root, ObjectNode params) {
            String id = params.path("templateId").asText();
            if (id.isBlank()) throw ServiceException.validation("cellTemplate.remove requires templateId");
            ArrayNode templates = SnapshotMutationSupport.array(root, "cellStyleTemplates");
            for (int index = templates.size() - 1; index >= 0; index--) {
                if (id.equals(templates.get(index).path("id").asText())) templates.remove(index);
            }
        }

        private void setEditor(ObjectNode root, String sheetId, ObjectNode params) {
            ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
            JsonNode editor = params.get("editor");
            if (editor != null && !editor.isNull() && !editor.isObject()) throw ServiceException.validation("cell.editor.set editor must be an object or null");
            for (RangeRef range : SnapshotMutationSupport.styleRanges(root, sheetId, params)) {
                for (int row = range.startRow(); row <= range.endRow(); row++) {
                    for (int column = range.startColumn(); column <= range.endColumn(); column++) {
                        ObjectNode cell = SnapshotMutationSupport.cell(sheet, new SnapshotMutationSupport.CellCoordinate(row, column), true);
                        if (editor == null || editor.isNull()) cell.remove("editor");
                        else cell.set("editor", editor.deepCopy());
                    }
                }
            }
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
            if (name.isBlank() || name.length() > com.xc.luckysheet.server.contract.GeneratedWorkbookContract.MAX_WORKBOOK_NAME_LENGTH) {
                throw ServiceException.validation("Workbook name is invalid");
            }
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
