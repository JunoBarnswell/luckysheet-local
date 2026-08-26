package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.service.ServiceException;

/**
 * Server-side boundary for the client-resolved data-region contract.
 *
 * The client may resolve the interaction context for latency, but the server
 * verifies the declared range, owner and header semantics against the current
 * snapshot before accepting a structural mutation.  There is no server-side
 * fallback to an old range or header convention.
 */
public final class DataRegionContextValidator {
    private DataRegionContextValidator() {
    }

    public static void validateSort(ObjectNode root, String sheetId, ObjectNode params) {
        ObjectNode context = requiredContext(params);
        RangeRef contextRange = range(context.get("range"), sheetId, "DataRegionContext range");
        RangeRef operationRange = range(params.get("range"), sheetId, "Sort range");
        JsonNode sortState = params.get("sortState");
        if (sortState != null && sortState.isObject() && sortState.has("range")) {
            RangeRef stateRange = range(sortState.get("range"), sheetId, "Sort state range");
            if (!sameRange(contextRange, stateRange)) throw ServiceException.validation("Sort DataRegionContext range does not match sort state");
        }
        if (operationRange.startRow() < contextRange.startRow() || operationRange.endRow() > contextRange.endRow()
                || operationRange.startColumn() < contextRange.startColumn() || operationRange.endColumn() > contextRange.endColumn()) {
            throw ServiceException.validation("Sort range is outside DataRegionContext range");
        }
        validateOwnerAndHeader(root, sheetId, context, contextRange);
        JsonNode hasHeader = params.get("hasHeader");
        JsonNode header = context.get("header");
        if (hasHeader != null && (!hasHeader.isBoolean() || hasHeader.asBoolean() != "present".equals(header.path("kind").asText()))) {
            throw ServiceException.validation("Sort header flag does not match DataRegionContext");
        }
    }

    public static void validateFilter(ObjectNode root, String sheetId, ObjectNode params, RangeRef filterRange, String ownerKind, String tableId) {
        ObjectNode context = requiredContext(params);
        RangeRef contextRange = range(context.get("range"), sheetId, "DataRegionContext range");
        if (!sameRange(contextRange, filterRange)) throw ServiceException.validation("Filter range does not match DataRegionContext");
        ObjectNode owner = (ObjectNode) context.get("owner");
        if (!ownerKind.equals(owner.path("kind").asText()) || (tableId == null ? owner.has("tableId") : !tableId.equals(owner.path("tableId").asText()))) {
            throw ServiceException.validation("Filter owner does not match DataRegionContext");
        }
        validateOwnerAndHeader(root, sheetId, context, contextRange);
    }

    private static ObjectNode requiredContext(ObjectNode params) {
        JsonNode raw = params.get("dataRegionContext");
        if (raw == null || !raw.isObject()) throw ServiceException.validation("DataRegionContext is required");
        ObjectNode context = (ObjectNode) raw;
        if (!"DataRegionContext".equals(context.path("schema").asText()) || context.path("version").asInt(-1) != 1) {
            throw ServiceException.validation("Unsupported DataRegionContext version");
        }
        if (!context.path("selection").isObject() || !context.path("currentRegion").isObject()
                || !context.path("usedRange").isObject() || !context.path("owner").isObject()
                || !context.path("header").isObject() || !context.path("visibleRows").isArray()) {
            throw ServiceException.validation("DataRegionContext is incomplete");
        }
        return context;
    }

    private static void validateOwnerAndHeader(ObjectNode root, String sheetId, ObjectNode context, RangeRef range) {
        ObjectNode sheet = findSheet(root, sheetId);
        ObjectNode owner = (ObjectNode) context.get("owner");
        String kind = owner.path("kind").asText();
        ObjectNode header = (ObjectNode) context.get("header");
        String headerKind = header.path("kind").asText();
        if (!"present".equals(headerKind) && !"absent".equals(headerKind)) throw ServiceException.validation("DataRegionContext header kind is invalid");
        if ("sheet-table".equals(kind)) {
            String tableId = owner.path("tableId").asText();
            ObjectNode table = findTable(sheet, tableId);
            RangeRef tableRange = range(table.get("range"), sheetId, "Sheet Table range");
            if (!sameRange(tableRange, range)) throw ServiceException.validation("DataRegionContext table owner does not own range");
            boolean hasHeader = table.path("hasHeaderRow").asBoolean();
            if (hasHeader != "present".equals(headerKind)) throw ServiceException.validation("DataRegionContext table header is stale");
            if (hasHeader && header.path("row").asInt(-1) != range.startRow()) throw ServiceException.validation("DataRegionContext table header row is invalid");
        } else if ("worksheet".equals(kind)) {
            if (owner.has("tableId")) throw ServiceException.validation("Worksheet DataRegionContext cannot carry tableId");
            for (JsonNode raw : sheet.path("sheetTables")) {
                if (!raw.isObject()) throw ServiceException.validation("Sheet Table is invalid");
                RangeRef tableRange = range(raw.get("range"), sheetId, "Sheet Table range");
                if (contains(tableRange, range.startRow(), range.startColumn())) throw ServiceException.validation("Worksheet DataRegionContext points into a Sheet Table");
            }
            boolean inferredHeader = inferWorksheetHeader(sheet, range);
            if (inferredHeader != "present".equals(headerKind)) throw ServiceException.validation("DataRegionContext worksheet header is stale");
            if (inferredHeader && header.path("row").asInt(-1) != range.startRow()) throw ServiceException.validation("Worksheet DataRegionContext header row is invalid");
        } else {
            throw ServiceException.validation("DataRegionContext owner kind is invalid");
        }
    }

    private static ObjectNode findSheet(ObjectNode root, String sheetId) {
        for (JsonNode raw : root.path("sheets")) if (raw.isObject() && sheetId.equals(raw.path("id").asText())) return (ObjectNode) raw;
        throw ServiceException.notFound("Sheet not found: " + sheetId);
    }

    private static ObjectNode findTable(ObjectNode sheet, String tableId) {
        for (JsonNode raw : sheet.path("sheetTables")) if (raw.isObject() && tableId.equals(raw.path("id").asText())) return (ObjectNode) raw;
        throw ServiceException.notFound("Sheet Table not found: " + tableId);
    }

    private static RangeRef range(JsonNode raw, String sheetId, String label) {
        if (raw == null || !raw.isObject() || !sheetId.equals(raw.path("sheetId").asText())) throw ServiceException.validation(label + " is invalid");
        try {
            return new RangeRef(sheetId, raw.path("startRow").asInt(-1), raw.path("endRow").asInt(-1), raw.path("startColumn").asInt(-1), raw.path("endColumn").asInt(-1));
        } catch (IllegalArgumentException error) {
            throw ServiceException.validation(label + " is invalid");
        }
    }

    private static boolean sameRange(RangeRef left, RangeRef right) {
        return left.sheetId().equals(right.sheetId()) && left.startRow() == right.startRow() && left.endRow() == right.endRow()
                && left.startColumn() == right.startColumn() && left.endColumn() == right.endColumn();
    }

    private static boolean contains(RangeRef range, int row, int column) {
        return row >= range.startRow() && row <= range.endRow() && column >= range.startColumn() && column <= range.endColumn();
    }

    private static boolean inferWorksheetHeader(ObjectNode sheet, RangeRef range) {
        JsonNode row = sheet.path("cells").path(Integer.toString(range.startRow()));
        int populated = 0;
        int textual = 0;
        for (int column = range.startColumn(); column <= range.endColumn(); column++) {
            JsonNode cell = row.path(Integer.toString(column));
            JsonNode value = cell.get("value");
            if (value == null || value.isNull() || (value.isTextual() && value.asText().isEmpty())) continue;
            populated++;
            if (value.isTextual()) textual++;
        }
        return populated > 0 && populated == textual;
    }
}
