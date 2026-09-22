package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/** Deterministically materializes and removes Pivot drill-down detail sheets. */
final class PivotDrillDownMutationDescriptor extends CanonicalJsonMutationDescriptor {
    private static final Set<String> ADD_KEYS = Set.of("sheetId", "pivotId", "label", "sourceRowPaths", "targetSheetId", "target");
    private static final Set<String> REMOVE_KEYS = Set.of("sheetId", "targetSheetId");
    static final Set<String> IDS = Set.of("pivot.drilldown.add", "pivot.drilldown.remove");

    PivotDrillDownMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR);
        if (!IDS.contains(id)) throw new IllegalArgumentException("Unsupported pivot drill-down mutation: " + id);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = PivotMutationDescriptor.canonicalSnapshot(snapshot);
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        if (id().equals("pivot.drilldown.remove")) {
            String targetSheetId = targetSheetId(params);
            return List.of(SnapshotMutationSupport.wholeSheetRange(root, targetSheetId));
        }
        DrillPlan plan = plan(root, mutation.sheetId(), params);
        List<RangeRef> ranges = new ArrayList<>(plan.sourceRanges());
        ranges.add(SnapshotMutationSupport.wholeSheetRange(root, mutation.sheetId()));
        // The target sheet is created by this same atomic mutation, so it has
        // no pre-existing protection state to resolve during prepare.
        return ranges;
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = PivotMutationDescriptor.canonicalSnapshot(snapshot);
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        if (id().equals("pivot.drilldown.remove")) {
            remove(root, targetSheetId(params));
            return root;
        }
        DrillPlan plan = plan(root, mutation.sheetId(), params);
        if (sheetExists(root, plan.targetSheetId())) throw ServiceException.conflict("Pivot drill-down target already exists: " + plan.targetSheetId());
        ObjectNode target = createSheet(plan.targetSheetId(), plan.sheetName(), plan.rowCount(), plan.columnCount());
        writePlan(root, target, plan);
        SnapshotMutationSupport.sheets(root).add(target);
        return root;
    }

    private DrillPlan plan(ObjectNode root, String pivotSheetId, ObjectNode params) {
        SnapshotMutationSupport.validateKnownKeys(params, ADD_KEYS, "pivot.drilldown.add params");
        String pivotId = SnapshotMutationSupport.text(params, "pivotId");
        ObjectNode pivotSheet = SnapshotMutationSupport.sheet(root, pivotSheetId);
        ObjectNode pivot = SnapshotMutationSupport.requireById(SnapshotMutationSupport.array(pivotSheet, "pivots"), pivotId, "Pivot");
        if (!"PivotDefinition".equals(SnapshotMutationSupport.text(pivot, "schema"))) throw ServiceException.validation("Pivot schema must be PivotDefinition");
        String label = SnapshotMutationSupport.text(params, "label");
        String targetSheetId = SnapshotMutationSupport.text(params, "targetSheetId");
        SnapshotMutationSupport.CellCoordinate anchor = coordinate(params.get("target"));
        ArrayNode paths = SnapshotMutationSupport.requiredArray(params, "sourceRowPaths");
        if (paths.size() > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Pivot drill-down has too many source rows");
        List<RangeRef> sourceRanges = PivotMutationDescriptor.sourceRanges(root, pivot);
        ObjectNode source = (ObjectNode) pivot.get("source");
        if ("data-source".equals(source.path("kind").asText())) {
            ObjectNode manifest = SnapshotMutationSupport.requireById(SnapshotMutationSupport.dataModelArray(root, "sources"),
                    SnapshotMutationSupport.text(source, "dataSourceId"), "Data source");
            if (!manifest.path("blocks").isEmpty()) throw unsupportedBlocks();
        }
        if (sourceRanges.isEmpty()) throw ServiceException.validation("Pivot drill-down source has no worksheet range");
        Map<String, SourceNode> nodes = new LinkedHashMap<>();
        for (int index = 0; index < sourceRanges.size(); index++) {
            String sourceId = "worksheet-ranges".equals(source.path("kind").asText())
                    ? SnapshotMutationSupport.text((ObjectNode) source.path("ranges").get(index), "sourceId") : "__single-source__";
            RangeRef range = sourceRanges.get(index);
            for (JsonNode region : SnapshotMutationSupport.array(SnapshotMutationSupport.sheet(root, range.sheetId()), "dataRegions")) {
                RangeRef bounds = SnapshotMutationSupport.range(root, region.get("range"));
                if (range.startRow() <= bounds.endRow() && range.endRow() >= bounds.startRow()
                        && range.startColumn() <= bounds.endColumn() && range.endColumn() >= bounds.startColumn()) throw unsupportedBlocks();
            }
            nodes.put(sourceId, new SourceNode(sourceId, range));
        }
        List<DrillColumn> columns = columns(root, List.copyOf(nodes.values()));
        if (columns.isEmpty()) throw ServiceException.validation("Pivot drill-down source has no columns");
        boolean multiSource = nodes.size() > 1;
        Set<String> optional = new HashSet<>();
        for (JsonNode relationship : source.path("relationships")) {
            if ("left".equals(relationship.path("join").asText())) optional.add(relationship.path("right").path("sourceId").asText());
        }
        List<String> roots = nodes.keySet().stream().filter(key -> !optional.contains(key)).toList();
        if (multiSource && !optional.isEmpty() && roots.size() != 1) throw ServiceException.validation("Pivot drill-down source graph has no deterministic root");
        Map<String, Map<String, SourcePath>> records = new LinkedHashMap<>();
        for (JsonNode raw : paths) {
            if (!raw.isObject()) throw ServiceException.validation("Pivot drill-down source path must be an object");
            ObjectNode path = (ObjectNode) raw;
            SnapshotMutationSupport.validateKnownKeys(path, Set.of("sheetId", "row", "sourceId", "recordId"), "Pivot drill-down source path");
            String sheetId = SnapshotMutationSupport.text(path, "sheetId");
            String sourceId = path.has("sourceId") ? SnapshotMutationSupport.text(path, "sourceId") : null;
            String recordId = path.has("recordId") ? SnapshotMutationSupport.text(path, "recordId") : null;
            if (multiSource && (sourceId == null || recordId == null)) throw ServiceException.validation("Joined Pivot drill-down provenance requires sourceId and recordId");
            SourceNode node = multiSource ? nodes.get(sourceId) : nodes.values().iterator().next();
            if (node == null) throw ServiceException.validation("Pivot drill-down provenance references an unknown source: " + sourceId);
            JsonNode rowValue = path.get("row");
            if (rowValue == null || !rowValue.isIntegralNumber() || !rowValue.canConvertToInt()) throw ServiceException.validation("Pivot drill-down source row is invalid");
            int row = rowValue.intValue();
            if (!sheetId.equals(node.range().sheetId()) || row <= node.range().startRow() || row > node.range().endRow()) {
                throw ServiceException.validation("Pivot drill-down provenance is outside its declared source range");
            }
            String identity = recordId == null ? sheetId + ":" + row : recordId;
            Map<String, SourcePath> record = records.computeIfAbsent(identity, ignored -> new LinkedHashMap<>());
            if (record.putIfAbsent(node.sourceId(), new SourcePath(sheetId, row)) != null) throw ServiceException.validation("Pivot drill-down provenance repeats a source in record " + identity);
        }
        for (var record : records.entrySet()) {
            for (String required : roots) {
                if (!record.getValue().containsKey(required)) throw ServiceException.validation("Pivot drill-down provenance is incomplete for record " + record.getKey());
            }
        }
        int detailRows = records.size();
        long targetRowCount = (long) anchor.row() + detailRows + 1L;
        long targetColumnCount = (long) anchor.column() + columns.size();
        if (targetRowCount > (long) SnapshotMutationSupport.MAX_ROW + 1L
                || targetColumnCount > (long) SnapshotMutationSupport.MAX_COLUMN + 1L) {
            throw ServiceException.validation("Pivot drill-down target exceeds the new worksheet bounds");
        }
        String sheetName = ("Drill " + pivotId + " " + label).substring(0, Math.min(31, ("Drill " + pivotId + " " + label).length()));
        return new DrillPlan(targetSheetId, sheetName, anchor, sourceRanges, columns, records.values().stream().map(Map::copyOf).toList(),
                Math.max(1_000, Math.toIntExact(targetRowCount)), Math.max(26, Math.toIntExact(targetColumnCount)));
    }

    private ServiceException unsupportedBlocks() {
        return new ServiceException("UNSUPPORTED_FEATURE", 422,
                "Pivot drill-down requires canonical block reads for this data source; no detail sheet was created");
    }

    private String targetSheetId(ObjectNode params) {
        Set<String> allowed = id().equals("pivot.drilldown.remove") ? REMOVE_KEYS : ADD_KEYS;
        SnapshotMutationSupport.validateKnownKeys(params, allowed, id() + " params");
        return SnapshotMutationSupport.text(params, "targetSheetId");
    }

    private List<DrillColumn> columns(ObjectNode root, List<SourceNode> nodes) {
        List<DrillColumn> columns = new ArrayList<>();
        Set<String> labels = new java.util.HashSet<>();
        for (SourceNode node : nodes) {
            RangeRef range = node.range();
            ObjectNode sheet = SnapshotMutationSupport.sheet(root, range.sheetId());
            for (int column = range.startColumn(); column <= range.endColumn(); column++) {
                JsonNode cell = SnapshotMutationSupport.cell(sheet, new SnapshotMutationSupport.CellCoordinate(range.startRow(), column), false);
                JsonNode raw = cell == null ? null : cell.hasNonNull("formulaValue") ? cell.get("formulaValue") : cell.get("value");
                JsonNode header = cellScalar(raw).get("value");
                String base = header.isNull() ? null : header.asText();
                if (base == null || base.isBlank()) base = "Column " + (column - range.startColumn() + 1);
                String label = base;
                if (labels.contains(label) && nodes.size() > 1) label = sheet.path("name").asText(range.sheetId()) + "." + base;
                int suffix = 2;
                while (labels.contains(label)) label = base + " (" + suffix++ + ")";
                labels.add(label);
                columns.add(new DrillColumn(node.sourceId(), column, label));
            }
        }
        return List.copyOf(columns);
    }

    private void writePlan(ObjectNode root, ObjectNode target, DrillPlan plan) {
        for (int index = 0; index < plan.columns().size(); index++) {
            SnapshotMutationSupport.putCell(target, new SnapshotMutationSupport.CellCoordinate(plan.anchor().row(), plan.anchor().column() + index), cell(plan.columns().get(index).label()));
        }
        for (int rowOffset = 0; rowOffset < plan.records().size(); rowOffset++) {
            Map<String, SourcePath> record = plan.records().get(rowOffset);
            for (int columnOffset = 0; columnOffset < plan.columns().size(); columnOffset++) {
                DrillColumn column = plan.columns().get(columnOffset);
                SourcePath path = record.get(column.sourceId());
                JsonNode value = null;
                if (path != null) {
                    ObjectNode sourceSheet = SnapshotMutationSupport.sheet(root, path.sheetId());
                    ObjectNode sourceCell = SnapshotMutationSupport.cell(sourceSheet, new SnapshotMutationSupport.CellCoordinate(path.row(), column.column()), false);
                    if (sourceCell != null) value = sourceCell.hasNonNull("formulaValue") ? sourceCell.get("formulaValue") : sourceCell.get("value");
                }
                SnapshotMutationSupport.putCell(target, new SnapshotMutationSupport.CellCoordinate(plan.anchor().row() + rowOffset + 1, plan.anchor().column() + columnOffset), cellScalar(value));
            }
        }
    }

    private void remove(ObjectNode root, String sheetId) {
        ArrayNode sheets = SnapshotMutationSupport.sheets(root);
        if (sheets.size() <= 1) throw ServiceException.validation("A workbook must keep at least one worksheet");
        for (int index = 0; index < sheets.size(); index++) {
            if (sheetId.equals(sheets.get(index).path("id").asText())) {
                if (!sheets.get(index).path("name").asText().startsWith("Drill ")) {
                    throw ServiceException.forbidden("Only a server-created pivot drill-down sheet may be removed by this mutation");
                }
                sheets.remove(index);
                removeScopedState(root, sheetId);
                return;
            }
        }
        throw ServiceException.notFound("Pivot drill-down target not found: " + sheetId);
    }

    private void removeScopedState(ObjectNode root, String sheetId) {
        ArrayNode documents = SnapshotMutationSupport.array(root, "printDocuments");
        for (int index = documents.size() - 1; index >= 0; index--) if (sheetId.equals(documents.get(index).path("sheetId").asText())) documents.remove(index);
        ArrayNode names = SnapshotMutationSupport.array(root, "definedNameModels");
        for (int index = names.size() - 1; index >= 0; index--) {
            JsonNode name = names.get(index);
            if (name.isObject() && "sheet".equals(name.path("scope").asText()) && sheetId.equals(name.path("sheetId").asText())) names.remove(index);
        }
    }

    private boolean sheetExists(ObjectNode root, String sheetId) {
        for (JsonNode sheet : SnapshotMutationSupport.sheets(root)) if (sheetId.equals(sheet.path("id").asText())) return true;
        return false;
    }

    private ObjectNode createSheet(String id, String name, int rowCount, int columnCount) {
        ObjectNode sheet = JsonNodeFactory.instance.objectNode();
        sheet.put("kind", "worksheet");
        sheet.put("id", id);
        sheet.put("name", name);
        sheet.put("rowCount", rowCount);
        sheet.put("columnCount", columnCount);
        sheet.set("cells", JsonNodeFactory.instance.objectNode());
        sheet.set("merges", JsonNodeFactory.instance.arrayNode());
        sheet.putObject("pane").put("kind", "none");
        sheet.put("defaultRowHeightPx", 20);
        sheet.put("defaultColumnWidthPx", 64);
        for (String property : List.of("pivots", "sparklines", "sparklineGroups", "drawings", "conditionalFormats", "dataValidations", "hiddenRows", "hiddenColumns", "sheetTables", "protectionRules")) sheet.set(property, JsonNodeFactory.instance.arrayNode());
        ObjectNode review = sheet.putObject("review");
        review.putObject("notesByCell");
        review.putObject("notesById");
        review.putObject("threadIdsByCell");
        review.putObject("threadsById");
        sheet.set("drawingPayloads", JsonNodeFactory.instance.objectNode());
        sheet.set("rowHeightsPx", JsonNodeFactory.instance.objectNode());
        sheet.set("columnWidthsPx", JsonNodeFactory.instance.objectNode());
        sheet.put("showGridlines", true);
        sheet.put("showHeaders", true);
        sheet.put("zoom", 100);
        sheet.put("hidden", false);
        sheet.putObject("outline").set("groups", JsonNodeFactory.instance.arrayNode());
        return sheet;
    }

    private SnapshotMutationSupport.CellCoordinate coordinate(JsonNode raw) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Pivot drill-down target anchor is invalid");
        JsonNode row = raw.get("row");
        JsonNode column = raw.get("column");
        if (row == null || !row.isIntegralNumber() || !row.canConvertToInt() || row.intValue() < 0
                || column == null || !column.isIntegralNumber() || !column.canConvertToInt() || column.intValue() < 0) throw ServiceException.validation("Pivot drill-down target anchor is invalid");
        return new SnapshotMutationSupport.CellCoordinate(row.intValue(), column.intValue());
    }

    private ObjectNode cell(String value) {
        ObjectNode cell = JsonNodeFactory.instance.objectNode();
        cell.put("value", value);
        return cell;
    }

    private ObjectNode cellScalar(JsonNode raw) {
        ObjectNode cell = JsonNodeFactory.instance.objectNode();
        if (raw == null || raw.isNull()) cell.putNull("value");
        else if (raw.isObject() && "error".equals(raw.path("kind").asText()) && raw.path("code").isTextual()) cell.put("value", raw.path("code").asText());
        else if (!raw.isTextual() && !raw.isNumber() && !raw.isBoolean()) throw new ServiceException("UNSUPPORTED_FEATURE", 422,
                "Pivot drill-down cannot represent this source value; no detail sheet was created");
        else cell.set("value", raw.deepCopy());
        return cell;
    }

    private record SourcePath(String sheetId, int row) {
    }

    private record SourceNode(String sourceId, RangeRef range) {
    }

    private record DrillColumn(String sourceId, int column, String label) {
    }

    private record DrillPlan(
            String targetSheetId,
            String sheetName,
            SnapshotMutationSupport.CellCoordinate anchor,
            List<RangeRef> sourceRanges,
            List<DrillColumn> columns,
            List<Map<String, SourcePath>> records,
            int rowCount,
            int columnCount
    ) {
    }
}
