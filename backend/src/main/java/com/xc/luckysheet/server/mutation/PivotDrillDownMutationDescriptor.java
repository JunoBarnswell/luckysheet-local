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
    private static final Set<String> ADD_KEYS = Set.of("sheetId", "pivotId", "label", "sourceRowPaths", "targetSheetId", "target", "detail");
    private static final Set<String> REMOVE_KEYS = Set.of("sheetId", "targetSheetId", "sourceId", "regionId");
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
            remove(root, params);
            return root;
        }
        DrillPlan plan = plan(root, mutation.sheetId(), params);
        if (sheetExists(root, plan.targetSheetId())) throw ServiceException.conflict("Pivot drill-down target already exists: " + plan.targetSheetId());
        ObjectNode target = createSheet(plan.targetSheetId(), plan.sheetName(), plan.rowCount(), plan.columnCount());
        writePlan(target, plan);
        SnapshotMutationSupport.sheets(root).add(target);
        SnapshotMutationSupport.dataModelArray(root, "sources").add(plan.detail().source().deepCopy());
        SnapshotMutationSupport.array(target, "dataRegions").add(plan.detail().region().deepCopy());
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
        if (paths.isEmpty()) throw ServiceException.validation("Pivot drill-down requires at least one source row");
        if (paths.size() > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Pivot drill-down has too many source rows");
        List<RangeRef> sourceRanges = PivotMutationDescriptor.sourceRanges(root, pivot);
        ObjectNode source = (ObjectNode) pivot.get("source");
        if (sourceRanges.isEmpty()) throw ServiceException.validation("Pivot drill-down source has no worksheet range");
        Map<String, SourceNode> nodes = new LinkedHashMap<>();
        for (int index = 0; index < sourceRanges.size(); index++) {
            String sourceId = "worksheet-ranges".equals(source.path("kind").asText())
                    ? SnapshotMutationSupport.text((ObjectNode) source.path("ranges").get(index), "sourceId") : "__single-source__";
            RangeRef range = sourceRanges.get(index);
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
        int rowCount = Math.max(1_000, Math.toIntExact(targetRowCount));
        int columnCount = Math.max(26, Math.toIntExact(targetColumnCount));
        MaterializedDetail detail = detail(root, params, targetSheetId, sheetName, anchor, columns, detailRows, rowCount, columnCount);
        return new DrillPlan(targetSheetId, sheetName, anchor, sourceRanges, columns, records.values().stream().map(Map::copyOf).toList(), detail,
                rowCount, columnCount);
    }

    private MaterializedDetail detail(ObjectNode root, ObjectNode params, String targetSheetId, String sheetName,
                                      SnapshotMutationSupport.CellCoordinate anchor, List<DrillColumn> columns,
                                      int detailRows, int rowCount, int columnCount) {
        JsonNode raw = params.get("detail");
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Pivot drill-down detail is required");
        ObjectNode detail = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(detail, Set.of("source", "region", "headers"), "Pivot drill-down detail");
        if (!detail.path("source").isObject() || !detail.path("region").isObject()) throw ServiceException.validation("Pivot drill-down source and region are required");
        ObjectNode source = (ObjectNode) detail.path("source");
        ObjectNode region = (ObjectNode) detail.path("region");
        ArrayNode headers = SnapshotMutationSupport.requiredArray(detail, "headers");
        if (headers.size() != columns.size()) throw ServiceException.validation("Pivot drill-down materialized columns do not match source columns");
        List<String> values = new ArrayList<>(headers.size());
        for (int index = 0; index < headers.size(); index++) {
            JsonNode header = headers.get(index);
            if (!header.isTextual() || !header.asText().equals(columns.get(index).label())) {
                throw ServiceException.validation("Pivot drill-down materialized header does not match source column");
            }
            values.add(header.asText());
        }
        ObjectNode validationRoot = root.deepCopy();
        ObjectNode validationSheet = createSheet(targetSheetId, sheetName, rowCount, columnCount);
        SnapshotMutationSupport.sheets(validationRoot).add(validationSheet);
        ObjectNode validatedSource = DataSourceMutationDescriptor.validateQuerySource(validationRoot, targetSheetId, source);
        if (SnapshotMutationSupport.findById(SnapshotMutationSupport.dataModelArray(root, "sources"), validatedSource.path("id").asText()) != null) {
            throw ServiceException.conflict("Pivot drill-down data source already exists: " + validatedSource.path("id").asText());
        }
        SnapshotMutationSupport.dataModelArray(validationRoot, "sources").add(validatedSource.deepCopy());
        ObjectNode validatedRegion = DataSourceMutationDescriptor.validateCompositeRegion(validationRoot, targetSheetId, region);
        RangeRef sourceRange = SnapshotMutationSupport.range(validationRoot, validatedSource.get("sourceRange"));
        RangeRef regionRange = SnapshotMutationSupport.range(validationRoot, validatedRegion.get("range"));
        int expectedEndRow = anchor.row() + detailRows;
        int expectedEndColumn = anchor.column() + columns.size() - 1;
        if (!"chunked-table".equals(validatedSource.path("kind").asText())
                || validatedSource.path("rowCount").asLong(-1) != detailRows
                || validatedSource.path("fields").size() != columns.size()
                || !validatedSource.path("id").asText().equals(validatedRegion.path("sourceId").asText())
                || validatedSource.path("revision").asLong(-1) != validatedRegion.path("revision").asLong(-2)
                || !targetSheetId.equals(sourceRange.sheetId()) || !targetSheetId.equals(regionRange.sheetId())
                || sourceRange.startRow() != anchor.row() || sourceRange.endRow() != expectedEndRow
                || sourceRange.startColumn() != anchor.column() || sourceRange.endColumn() != expectedEndColumn
                || !sourceRange.equals(regionRange) || validatedRegion.path("headerRow").asInt(-1) != anchor.row()) {
            throw ServiceException.validation("Pivot drill-down detail source does not match its target or provenance");
        }
        for (int index = 0; index < columns.size(); index++) {
            if (!values.get(index).equals(validatedSource.path("fields").get(index).path("name").asText())) {
                throw ServiceException.validation("Pivot drill-down field name does not match its source column");
            }
        }
        return new MaterializedDetail(validatedSource.deepCopy(), validatedRegion.deepCopy(), List.copyOf(values));
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

    private void writePlan(ObjectNode target, DrillPlan plan) {
        for (int index = 0; index < plan.detail().headers().size(); index++) {
            SnapshotMutationSupport.putCell(target, new SnapshotMutationSupport.CellCoordinate(plan.anchor().row(), plan.anchor().column() + index), cell(plan.detail().headers().get(index)));
        }
    }

    private void remove(ObjectNode root, ObjectNode params) {
        SnapshotMutationSupport.validateKnownKeys(params, REMOVE_KEYS, "pivot.drilldown.remove params");
        String sheetId = SnapshotMutationSupport.text(params, "targetSheetId");
        String sourceId = SnapshotMutationSupport.text(params, "sourceId");
        String regionId = SnapshotMutationSupport.text(params, "regionId");
        ArrayNode sheets = SnapshotMutationSupport.sheets(root);
        if (sheets.size() <= 1) throw ServiceException.validation("A workbook must keep at least one worksheet");
        for (int index = 0; index < sheets.size(); index++) {
            if (sheetId.equals(sheets.get(index).path("id").asText())) {
                ArrayNode regions = SnapshotMutationSupport.array((ObjectNode) sheets.get(index), "dataRegions");
                ObjectNode region = SnapshotMutationSupport.requireById(regions, regionId, "Pivot drill-down data region");
                if (!sourceId.equals(region.path("sourceId").asText()) || regions.size() != 1) {
                    throw ServiceException.conflict("Pivot drill-down sheet contains a different data binding");
                }
                SnapshotMutationSupport.removeById(regions, regionId);
                ArrayNode sources = SnapshotMutationSupport.dataModelArray(root, "sources");
                SnapshotMutationSupport.requireById(sources, sourceId, "Pivot drill-down data source");
                for (JsonNode rawSheet : sheets) {
                    if (!rawSheet.isObject() || sheetId.equals(rawSheet.path("id").asText())) continue;
                    for (JsonNode otherRegion : SnapshotMutationSupport.array((ObjectNode) rawSheet, "dataRegions")) {
                        if (sourceId.equals(otherRegion.path("sourceId").asText())) {
                            throw ServiceException.conflict("Pivot drill-down data source is referenced by another sheet");
                        }
                    }
                    for (JsonNode otherPivot : SnapshotMutationSupport.array((ObjectNode) rawSheet, "pivots")) {
                        if ("data-source".equals(otherPivot.path("source").path("kind").asText())
                                && sourceId.equals(otherPivot.path("source").path("dataSourceId").asText())) {
                            throw ServiceException.conflict("Pivot drill-down data source is referenced by another PivotTable");
                        }
                    }
                }
                for (JsonNode table : SnapshotMutationSupport.dataModelArray(root, "tables")) {
                    if (sourceId.equals(table.path("sourceId").asText())) {
                        throw ServiceException.conflict("Pivot drill-down data source is referenced by a workbook table");
                    }
                }
                SnapshotMutationSupport.removeById(sources, sourceId);
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
        for (String property : List.of("pivots", "sparklines", "sparklineGroups", "drawings", "conditionalFormats", "dataValidations", "hiddenRows", "hiddenColumns", "sheetTables", "dataRegions", "protectionRules")) sheet.set(property, JsonNodeFactory.instance.arrayNode());
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
                "Pivot drill-down cannot represent this source header; no detail sheet was created");
        else cell.set("value", raw.deepCopy());
        return cell;
    }

    private record SourcePath(String sheetId, int row) {
    }

    private record SourceNode(String sourceId, RangeRef range) {
    }

    private record DrillColumn(String sourceId, int column, String label) {
    }

    private record MaterializedDetail(ObjectNode source, ObjectNode region, List<String> headers) {
    }

    private record DrillPlan(
            String targetSheetId,
            String sheetName,
            SnapshotMutationSupport.CellCoordinate anchor,
            List<RangeRef> sourceRanges,
            List<DrillColumn> columns,
            List<Map<String, SourcePath>> records,
            MaterializedDetail detail,
            int rowCount,
            int columnCount
    ) {
    }
}
