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
import java.util.List;
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
        ranges.add(plan.targetRange());
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
        ObjectNode target = createSheet(plan.targetSheetId(), plan.sheetName());
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
        if (sourceRanges.isEmpty()) throw ServiceException.validation("Pivot drill-down source has no worksheet range");
        List<DrillColumn> columns = columns(root, sourceRanges);
        if (columns.isEmpty()) throw ServiceException.validation("Pivot drill-down source has no columns");
        List<SourcePath> sourcePaths = new ArrayList<>();
        for (JsonNode raw : paths) {
            if (!raw.isObject()) throw ServiceException.validation("Pivot drill-down source path must be an object");
            ObjectNode path = (ObjectNode) raw;
            SnapshotMutationSupport.validateKnownKeys(path, Set.of("sheetId", "row"), "Pivot drill-down source path");
            String sheetId = SnapshotMutationSupport.text(path, "sheetId");
            ObjectNode coordinate = JsonNodeFactory.instance.objectNode();
            coordinate.set("row", path.get("row"));
            coordinate.put("column", 0);
            int row = SnapshotMutationSupport.index(root, sheetId, coordinate, "row");
            sourcePaths.add(new SourcePath(sheetId, row));
        }
        int rowsPerResult = Math.max(sourceRanges.size(), 1);
        int detailRows = (int) Math.ceil(sourcePaths.size() / (double) rowsPerResult);
        if (anchor.row() + detailRows >= 1_000 || anchor.column() + columns.size() - 1 >= 26) {
            throw ServiceException.validation("Pivot drill-down target exceeds the new worksheet bounds");
        }
        RangeRef targetRange = new RangeRef(targetSheetId, anchor.row(), anchor.row() + detailRows, anchor.column(), anchor.column() + columns.size() - 1);
        String sheetName = ("Drill " + pivotId + " " + label).substring(0, Math.min(31, ("Drill " + pivotId + " " + label).length()));
        return new DrillPlan(targetSheetId, sheetName, anchor, sourceRanges, columns, sourcePaths, targetRange, rowsPerResult);
    }

    private String targetSheetId(ObjectNode params) {
        Set<String> allowed = id().equals("pivot.drilldown.remove") ? REMOVE_KEYS : ADD_KEYS;
        SnapshotMutationSupport.validateKnownKeys(params, allowed, id() + " params");
        return SnapshotMutationSupport.text(params, "targetSheetId");
    }

    private List<DrillColumn> columns(ObjectNode root, List<RangeRef> ranges) {
        List<DrillColumn> columns = new ArrayList<>();
        Set<String> labels = new java.util.HashSet<>();
        for (RangeRef range : ranges) {
            ObjectNode sheet = SnapshotMutationSupport.sheet(root, range.sheetId());
            for (int column = range.startColumn(); column <= range.endColumn(); column++) {
                JsonNode cell = SnapshotMutationSupport.cell(sheet, new SnapshotMutationSupport.CellCoordinate(range.startRow(), column), false);
                String base = scalar(cell == null ? null : cell.get("value"));
                if (base == null || base.isBlank()) base = "Column " + (column - range.startColumn() + 1);
                String label = base;
                if (labels.contains(label) && ranges.size() > 1) label = sheet.path("name").asText(range.sheetId()) + "." + base;
                int suffix = 2;
                while (labels.contains(label)) label = base + " (" + suffix++ + ")";
                labels.add(label);
                columns.add(new DrillColumn(range, column, label));
            }
        }
        return List.copyOf(columns);
    }

    private void writePlan(ObjectNode root, ObjectNode target, DrillPlan plan) {
        for (int index = 0; index < plan.columns().size(); index++) {
            SnapshotMutationSupport.putCell(target, new SnapshotMutationSupport.CellCoordinate(plan.anchor().row(), plan.anchor().column() + index), cell(plan.columns().get(index).label()));
        }
        int resultRows = (int) Math.ceil(plan.paths().size() / (double) plan.rowsPerResult());
        for (int rowOffset = 0; rowOffset < resultRows; rowOffset++) {
            List<SourcePath> paths = plan.paths().subList(rowOffset * plan.rowsPerResult(), Math.min(plan.paths().size(), (rowOffset + 1) * plan.rowsPerResult()));
            for (int columnOffset = 0; columnOffset < plan.columns().size(); columnOffset++) {
                DrillColumn column = plan.columns().get(columnOffset);
                SourcePath path = paths.stream().filter(candidate -> candidate.sheetId().equals(column.range().sheetId())).findFirst().orElse(null);
                JsonNode value = null;
                if (path != null) {
                    ObjectNode sourceSheet = SnapshotMutationSupport.sheet(root, path.sheetId());
                    ObjectNode sourceCell = SnapshotMutationSupport.cell(sourceSheet, new SnapshotMutationSupport.CellCoordinate(path.row(), column.column()), false);
                    if (sourceCell != null) value = sourceCell.has("formulaValue") ? sourceCell.get("formulaValue") : sourceCell.get("value");
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

    private ObjectNode createSheet(String id, String name) {
        ObjectNode sheet = JsonNodeFactory.instance.objectNode();
        sheet.put("id", id);
        sheet.put("name", name);
        sheet.put("rowCount", 1_000);
        sheet.put("columnCount", 26);
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
        if (row == null || !row.isIntegralNumber() || row.intValue() < 0 || column == null || !column.isIntegralNumber() || column.intValue() < 0) throw ServiceException.validation("Pivot drill-down target anchor is invalid");
        return new SnapshotMutationSupport.CellCoordinate(row.intValue(), column.intValue());
    }

    private ObjectNode cell(String value) {
        ObjectNode cell = JsonNodeFactory.instance.objectNode();
        cell.put("value", value);
        return cell;
    }

    private ObjectNode cellScalar(JsonNode raw) {
        ObjectNode cell = JsonNodeFactory.instance.objectNode();
        if (raw == null || raw.isNull() || (!raw.isTextual() && !raw.isNumber() && !raw.isBoolean())) cell.putNull("value");
        else cell.set("value", raw.deepCopy());
        return cell;
    }

    private String scalar(JsonNode raw) {
        if (raw == null || raw.isNull()) return null;
        if (raw.isTextual()) return raw.asText();
        if (raw.isNumber()) return raw.asText();
        if (raw.isBoolean()) return Boolean.toString(raw.asBoolean());
        return null;
    }

    private record SourcePath(String sheetId, int row) {
    }

    private record DrillColumn(RangeRef range, int column, String label) {
    }

    private record DrillPlan(
            String targetSheetId,
            String sheetName,
            SnapshotMutationSupport.CellCoordinate anchor,
            List<RangeRef> sourceRanges,
            List<DrillColumn> columns,
            List<SourcePath> paths,
            RangeRef targetRange,
            int rowsPerResult
    ) {
    }
}
