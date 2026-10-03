package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookRole;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * Server reducer for the canonical fill command.
 *
 * The browser sends the side-effect-free planner output as before/after cell
 * snapshots. The server derives the affected target band, validates the
 * one-axis geometry and verifies every before image while the workbook
 * operation lock is held. It independently derives copy, numeric and calendar
 * results and rejects missing, extra or forged writes before committing.
 */
public final class FillMutationDescriptor extends CanonicalJsonMutationDescriptor {
    public static final Set<String> IDS = Set.of("fill.applied", "fill.restored");

    public FillMutationDescriptor(String id) {
        super(id, WorkbookRole.EDITOR);
        if (!IDS.contains(id)) throw new IllegalArgumentException("Unknown fill mutation: " + id);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        return List.of(validatePayload(root, mutation.sheetId(), params));
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        RangeRef target = validatePayload(root, mutation.sheetId(), params);
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, mutation.sheetId());
        ArrayNode writes = (ArrayNode) params.get("writes");
        if ("fill.applied".equals(id())) assertCanonicalPlan(root, sheet, mutation.sheetId(), params);

        // Validate all snapshots before applying the first write. The reducer
        // itself is pure over a copied snapshot, but this also makes the
        // fail-close boundary explicit for a malformed mixed payload.
        for (JsonNode value : writes) {
            ObjectNode entry = requireWrite(value);
            SnapshotMutationSupport.CellCoordinate coordinate = writeCoordinate(root, mutation.sheetId(), target, entry);
            JsonNode current = SnapshotMutationSupport.cell(sheet, coordinate, false);
            assertBefore(current, entry.get("before"), coordinate);
            assertAfter(entry.get("after"));
        }
        for (JsonNode value : writes) {
            ObjectNode entry = (ObjectNode) value;
            SnapshotMutationSupport.CellCoordinate coordinate = writeCoordinate(root, mutation.sheetId(), target, entry);
            JsonNode after = entry.get("after");
            if (after == null || after.isNull()) SnapshotMutationSupport.removeCell(sheet, coordinate);
            else SnapshotMutationSupport.putCell(sheet, coordinate, after);
        }
        return root;
    }

    private static boolean sameJson(JsonNode left, JsonNode right) {
        if (left == null || left.isNull()) return right == null || right.isNull();
        if (right == null) return false;
        if (left.isNumber() && right.isNumber()) return left.asDouble() == right.asDouble();
        if (left.isObject() && right.isObject() && left.size() == right.size()) { for (var names = left.fieldNames(); names.hasNext();) { String name = names.next(); if (!sameJson(left.get(name), right.get(name))) return false; } return true; }
        if (left.isArray() && right.isArray() && left.size() == right.size()) { for (int index = 0; index < left.size(); index++) if (!sameJson(left.get(index), right.get(index))) return false; return true; }
        return left.equals(right);
    }
    private static boolean isDateFormat(String format) {
        boolean quoted = false;
        for (int i = 0; i < format.length(); i++) {
            char c = Character.toLowerCase(format.charAt(i));
            if (c == '"') { quoted = !quoted; continue; }
            if (quoted) continue;
            if (c == '\\') { i++; continue; }
            if (c == '[') {
                int end = format.indexOf(']', i + 1);
                if (end < 0) return false;
                i = end; continue;
            }
            if (c == 'y' || c == 'm' || c == 'd' || c == 'h' || c == 's') return true;
        }
        return false;
    }
    private record Seed(ObjectNode cell, double value, int travel, int row, int column, boolean date) { }
    private void assertCanonicalPlan(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
        RangeRef source = requireOwnRange(root, sheetId, params, "sourceRange");
        RangeRef target = requireOwnRange(root, sheetId, params, "targetRange");
        String direction = params.path("direction").asText(); boolean rows = "down".equals(direction) || "up".equals(direction);
        boolean reverse = "up".equals(direction) || "left".equals(direction);
        java.util.Map<String, JsonNode> writes = new java.util.HashMap<>();
        for (JsonNode write : params.path("writes")) writes.put(write.path("row").asInt() + ":" + write.path("column").asInt(), write.get("after"));
        java.util.Map<Integer, java.util.List<Seed>> tracks = new java.util.HashMap<>();
        java.util.Set<String> seeds = new java.util.HashSet<>();
        boolean series = "series".equals(params.path("mode").asText()) && !"autofill".equals(params.path("series").path("type").asText()); JsonNode options = params.path("series");
        String dateSystem = params.path("dateSystem").asText("1900");
        if (!java.util.Set.of("1900", "1904").contains(dateSystem)) throw ServiceException.validation("Invalid fill date system");
        if (series) for (int row = source.startRow(); row <= source.endRow(); row++) for (int column = source.startColumn(); column <= source.endColumn(); column++) {
            JsonNode cell = SnapshotMutationSupport.cell(sheet, new SnapshotMutationSupport.CellCoordinate(row, column), false);
            if (cell == null || cell.path("value").isNull() || cell.path("value").isMissingNode() || cell.path("value").isTextual() && cell.path("value").asText().isEmpty()) continue;
            if (cell.hasNonNull("formula") || cell.hasNonNull("formulaValue")) throw ServiceException.validation("Series fill requires numeric seeds");
            boolean date = "date".equals(options.path("type").asText()) || isDateFormat(cell.path("numberFormat").asText());
            double value;
            if (date) value = com.xc.luckysheet.server.contract.CanonicalExcelDate.serial(cell.path("value"), dateSystem);
            else { if (!cell.path("value").isNumber()) throw ServiceException.validation("Series fill requires finite numeric seeds"); value = cell.path("value").asDouble(); }
            if (!Double.isFinite(value)) throw ServiceException.validation("Series seed must be finite");
            int travel = (rows ? row : column) * (reverse ? -1 : 1);
            tracks.computeIfAbsent(rows ? column : row, ignored -> new java.util.ArrayList<>()).add(new Seed((ObjectNode) cell, value, travel, row, column, date)); seeds.add(row + ":" + column);
        }
        for (var entries : tracks.values()) {
            entries.sort(java.util.Comparator.comparingInt(Seed::travel));
            if (entries.stream().anyMatch(seed -> seed.date() != entries.getFirst().date())) throw ServiceException.validation("Series track mixes dates and numbers");
        }
        java.util.Map<Integer, double[]> coefficients = new java.util.HashMap<>();
        String seriesType = options.path("type").asText("linear");
        if (series && !java.util.Set.of("growth", "linear", "autofill", "date").contains(seriesType)) throw ServiceException.validation("Unsupported series type");
        if (series && !"growth".equals(seriesType) && !"date".equals(seriesType)) for (var track : tracks.entrySet()) {
            var entries = track.getValue(); Seed first = entries.getFirst();
            double step = options.has("stepValue") ? options.path("stepValue").asDouble() : entries.size() < 2 ? 1 : (entries.get(1).value() - first.value()) / (entries.get(1).travel() - first.travel());
            double intercept = first.value() - step * first.travel();
            if (options.path("trend").asBoolean() && !options.has("stepValue") && entries.size() >= 2) {
                double meanTravel = 0, meanValue = 0; for (Seed seed : entries) { meanTravel += seed.travel(); meanValue += seed.value(); } meanTravel /= entries.size(); meanValue /= entries.size();
                double numerator = 0, denominator = 0; for (Seed seed : entries) { numerator += (seed.travel() - meanTravel) * (seed.value() - meanValue); denominator += Math.pow(seed.travel() - meanTravel, 2); }
                if (denominator == 0) throw ServiceException.validation("Trend series needs distinct seeds"); step = numerator / denominator; intercept = meanValue - step * meanTravel;
            } else if (!options.has("stepValue")) for (Seed seed : entries) { double predicted = first.value() + step * (seed.travel() - first.travel()); if (Math.abs(predicted - seed.value()) > Math.ulp(1.0) * Math.max(1, Math.max(Math.abs(predicted), Math.abs(seed.value()))) * 16) throw ServiceException.validation("Series seeds do not define one progression"); }
            coefficients.put(track.getKey(), new double[]{step, intercept});
        }
        int changed = 0;
        for (int row = target.startRow(); row <= target.endRow(); row++) for (int column = target.startColumn(); column <= target.endColumn(); column++) {
            String key = row + ":" + column; JsonNode expected;
            if (!series) {
                int height = source.endRow() - source.startRow() + 1, width = source.endColumn() - source.startColumn() + 1;
                int sr = "up".equals(direction) ? source.endRow() - Math.floorMod(source.endRow() - row, height) : source.startRow() + Math.floorMod(row - source.startRow(), height);
                int sc = "left".equals(direction) ? source.endColumn() - Math.floorMod(source.endColumn() - column, width) : source.startColumn() + Math.floorMod(column - source.startColumn(), width);
                JsonNode raw = SnapshotMutationSupport.cell(sheet, new SnapshotMutationSupport.CellCoordinate(sr, sc), false);
                expected = raw == null ? null : raw.deepCopy();
                if (expected instanceof ObjectNode cell) {
                    cell.remove("formulaMetadata");
                    if (cell.hasNonNull("formula")) {
                        cell.put("formula", FormulaReferenceTransformer.offsetForCopy(cell.path("formula").asText(), row - sr, column - sc)); cell.putNull("value"); cell.remove(java.util.List.of("formulaValue", "displayValue"));
                    }
                    JsonNode presentation = cell.path("presentation");
                    if ("barcode".equals(presentation.path("kind").asText()) && "formula".equals(presentation.path("source").path("kind").asText())) ((ObjectNode) presentation.path("source")).put("formula", FormulaReferenceTransformer.offsetForCopy(presentation.path("source").path("formula").asText(), row - sr, column - sc));
                }
            } else {
                if (seeds.contains(key)) { if (writes.containsKey(key)) throw ServiceException.validation("Series fill cannot rewrite seeds"); continue; }
                var entries = tracks.get(rows ? column : row); if (entries == null || entries.isEmpty()) throw ServiceException.validation("Series requires a seed on every affected track");
                Seed first = entries.getFirst(); double distance = (rows ? row : column) * (reverse ? -1 : 1) - first.travel();
                String type = options.path("type").asText("linear"); double value;
                if ("date".equals(type)) value = com.xc.luckysheet.server.contract.CanonicalExcelDate.shift(first.value(), options.path("stepValue").asDouble(1) * distance, options.path("dateUnit").asText("day"), dateSystem);
                else if ("growth".equals(type)) { double ratio = options.has("stepValue") ? options.path("stepValue").asDouble() : entries.size() >= 2 && first.value() != 0 ? entries.get(1).value() / first.value() : 2; value = first.value() * Math.pow(ratio, distance); }
                else {
                    if (!java.util.Set.of("linear", "autofill").contains(type)) throw ServiceException.validation("Unsupported series type");
                    double step = coefficients.get(rows ? column : row)[0], intercept = coefficients.get(rows ? column : row)[1];
                    value = intercept + step * (first.travel() + distance);
                }
                if (!Double.isFinite(value)) throw ServiceException.validation("Series result must be finite");
                if (options.has("stopValue") && (value - first.value() >= 0 ? value > options.path("stopValue").asDouble() : value < options.path("stopValue").asDouble())) continue;
                ObjectNode cell = first.cell().deepCopy(); if (first.date() || "date".equals(type)) cell.put("value", com.xc.luckysheet.server.contract.CanonicalExcelDate.iso(value, dateSystem)); else cell.put("value", value); cell.remove(java.util.List.of("formula", "formulaValue", "displayValue", "formulaMetadata")); expected = cell;
            }
            JsonNode current = SnapshotMutationSupport.cell(sheet, new SnapshotMutationSupport.CellCoordinate(row, column), false);
            if (!sameJson(current, expected)) { changed++; if (!writes.containsKey(key) || !sameJson(expected, writes.get(key))) throw ServiceException.validation("Fill does not match the server canonical plan"); }
            else if (writes.containsKey(key)) throw ServiceException.validation("Fill includes an unchanged cell");
        }
        if (changed != writes.size()) throw ServiceException.validation("Fill includes writes outside the canonical plan");
    }

    private RangeRef validatePayload(ObjectNode root, String sheetId, ObjectNode params) {
        if (!sheetId.equals(params.path("sheetId").asText())) throw ServiceException.validation("Fill sheetId does not match mutation sheetId");
        RangeRef source = requireOwnRange(root, sheetId, params, "sourceRange");
        RangeRef target = requireOwnRange(root, sheetId, params, "targetRange");
        String direction = SnapshotMutationSupport.text(params, "direction");
        String mode = SnapshotMutationSupport.text(params, "mode");
        if (!Set.of("down", "up", "right", "left").contains(direction)) throw ServiceException.validation("Fill direction is invalid");
        if (!Set.of("copy", "series").contains(mode)) throw ServiceException.validation("Fill mode is invalid");
        assertGeometry(source, target, direction);
        if (params.has("series")) {
            JsonNode options = params.get("series");
            if (!options.isObject()) throw ServiceException.validation("Fill series options must be an object");
            for (String key : java.util.List.of("stepValue", "stopValue")) if (options.has(key) && (!options.get(key).isNumber() || !Double.isFinite(options.get(key).asDouble()))) throw ServiceException.validation("Fill numeric option must be finite");
            if (options.has("trend") && !options.get("trend").isBoolean()) throw ServiceException.validation("Fill trend must be boolean");
            if (options.has("seriesIn") && !java.util.Set.of("rows", "columns").contains(options.path("seriesIn").asText())) throw ServiceException.validation("Fill series axis is invalid");
            if (options.has("dateUnit") && !java.util.Set.of("day", "weekday", "month", "year").contains(options.path("dateUnit").asText())) throw ServiceException.validation("Fill date unit is invalid");
        }
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        int rowCount = sheet.path("rowCount").asInt(-1);
        int columnCount = sheet.path("columnCount").asInt(-1);
        if (rowCount < 1 || columnCount < 1 || target.endRow() >= rowCount || target.endColumn() >= columnCount) {
            throw ServiceException.validation("Fill target is outside worksheet bounds");
        }
        JsonNode writesNode = params.get("writes");
        if (writesNode == null || !writesNode.isArray() || writesNode.isEmpty() || writesNode.size() > SnapshotMutationSupport.MAX_CHANGED_CELLS) {
            throw ServiceException.validation("Fill writes are required and bounded");
        }
        Set<String> coordinates = new HashSet<>();
        for (JsonNode value : writesNode) {
            ObjectNode entry = requireWrite(value);
            SnapshotMutationSupport.CellCoordinate coordinate = writeCoordinate(root, sheetId, target, entry);
            if (!coordinates.add(coordinate.row() + ":" + coordinate.column())) throw ServiceException.validation("Fill writes contain a duplicate coordinate");
            assertAfter(entry.get("after"));
        }
        return target;
    }

    private RangeRef requireOwnRange(ObjectNode root, String sheetId, ObjectNode params, String property) {
        RangeRef range = SnapshotMutationSupport.range(root, params.get(property));
        SnapshotMutationSupport.requireSheet(range, sheetId);
        if (SnapshotMutationSupport.cellCount(range) > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Fill range is too large");
        return range;
    }

    private void assertGeometry(RangeRef source, RangeRef target, String direction) {
        boolean contains = target.startRow() <= source.startRow() && target.endRow() >= source.endRow()
                && target.startColumn() <= source.startColumn() && target.endColumn() >= source.endColumn();
        if (!contains) throw ServiceException.validation("Fill target must contain source range");
        boolean sameColumns = source.startColumn() == target.startColumn() && source.endColumn() == target.endColumn();
        boolean sameRows = source.startRow() == target.startRow() && source.endRow() == target.endRow();
        boolean valid = switch (direction) {
            case "down" -> sameColumns && target.startRow() == source.startRow() && target.endRow() >= source.endRow();
            case "up" -> sameColumns && target.endRow() == source.endRow() && target.startRow() <= source.startRow();
            case "right" -> sameRows && target.startColumn() == source.startColumn() && target.endColumn() >= source.endColumn();
            case "left" -> sameRows && target.endColumn() == source.endColumn() && target.startColumn() <= source.startColumn();
            default -> false;
        };
        if (!valid) throw ServiceException.validation("Fill direction requires a contiguous one-axis target extension");
    }

    private ObjectNode requireWrite(JsonNode value) {
        if (value == null || !value.isObject()) throw ServiceException.validation("Fill write must be an object");
        ObjectNode entry = (ObjectNode) value;
        if (!entry.has("row") || !entry.path("row").isIntegralNumber()
                || !entry.has("column") || !entry.path("column").isIntegralNumber()) {
            throw ServiceException.validation("Fill write coordinates are required");
        }
        return entry;
    }

    private SnapshotMutationSupport.CellCoordinate writeCoordinate(
            ObjectNode root,
            String sheetId,
            RangeRef target,
            ObjectNode entry
    ) {
        int row = SnapshotMutationSupport.index(root, sheetId, entry, "row");
        int column = SnapshotMutationSupport.index(root, sheetId, entry, "column");
        SnapshotMutationSupport.CellCoordinate coordinate = new SnapshotMutationSupport.CellCoordinate(row, column);
        if (!SnapshotMutationSupport.contains(target, coordinate)) throw ServiceException.validation("Fill write is outside target range");
        return coordinate;
    }

    private void assertBefore(JsonNode current, JsonNode before, SnapshotMutationSupport.CellCoordinate coordinate) {
        if (before == null || before.isNull()) {
            if (current != null) throw ServiceException.conflict("Fill target changed at " + coordinate.row() + ":" + coordinate.column());
            return;
        }
        if (!before.isObject() || current == null || !current.equals(before)) {
            throw ServiceException.conflict("Fill target changed at " + coordinate.row() + ":" + coordinate.column());
        }
    }

    private void assertAfter(JsonNode after) {
        if (after == null || after.isNull()) return;
        if (!after.isObject() || !after.has("value")) throw ServiceException.validation("Fill after cell must contain a value");
    }
}
