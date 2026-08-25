package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.List;

/**
 * Resolves the single AutoFilter owner for every range in a worksheet.
 *
 * Worksheet and Table filters are deliberately exclusive: there is no
 * precedence rule that can make an overlapping pair canonical.  Every
 * mutation and every snapshot boundary therefore uses this validator to
 * reject the transition before it can be observed by history or peers.
 */
public final class AutoFilterOwnershipValidator {
    private AutoFilterOwnershipValidator() {
    }

    public record Owner(String kind, String tableId, RangeRef range) {
        public boolean isSameOwner(String candidateKind, String candidateTableId) {
            return kind.equals(candidateKind)
                    && (tableId == null ? candidateTableId == null : tableId.equals(candidateTableId));
        }
    }

    /** Validate all current Worksheet/Table filter owners on a sheet. */
    public static List<Owner> resolveOwners(ObjectNode sheet, String sheetId) {
        List<Owner> owners = new ArrayList<>();
        JsonNode worksheetFilter = sheet.get("autoFilter");
        if (worksheetFilter != null && !worksheetFilter.isNull()) {
            if (!worksheetFilter.isObject()) throw ServiceException.validation("Worksheet AutoFilter is invalid");
            owners.add(new Owner("worksheet", null, rangeOf(worksheetFilter.get("range"), sheetId, "Worksheet AutoFilter")));
        }

        JsonNode tables = sheet.get("sheetTables");
        if (tables != null && !tables.isNull()) {
            if (!tables.isArray()) throw ServiceException.validation("Workbook snapshot sheetTables is invalid");
            for (JsonNode rawTable : tables) {
                if (!rawTable.isObject()) throw ServiceException.validation("Sheet Table is invalid");
                ObjectNode table = (ObjectNode) rawTable;
                String tableId = table.path("id").asText();
                if (tableId.isBlank()) throw ServiceException.validation("Sheet Table id is required");
                RangeRef tableRange = rangeOf(table.get("range"), sheetId, "Sheet Table range");
                JsonNode tableFilter = table.get("autoFilter");
                if (tableFilter == null || tableFilter.isNull()) continue;
                if (!tableFilter.isObject()) throw ServiceException.validation("Table AutoFilter is invalid");
                RangeRef filterRange = rangeOf(tableFilter.get("range"), sheetId, "Table AutoFilter");
                if (!sameRange(filterRange, tableRange)) {
                    throw ServiceException.validation("Table AutoFilter range must equal the Table range");
                }
                owners.add(new Owner("table", tableId, filterRange));
            }
        }

        for (int left = 0; left < owners.size(); left++) {
            for (int right = left + 1; right < owners.size(); right++) {
                Owner first = owners.get(left);
                Owner second = owners.get(right);
                if (overlaps(first.range(), second.range())) {
                    throw ServiceException.validation("AutoFilter owners cannot overlap: " + ownerName(first) + " and " + ownerName(second));
                }
            }
        }
        return List.copyOf(owners);
    }

    /**
     * Validate a candidate owner against the current snapshot.  The existing
     * owner with the same identity is excluded so resize/update operations can
     * replace its range atomically.
     */
    public static void validateCandidate(ObjectNode sheet, String sheetId, String candidateKind, String candidateTableId, RangeRef candidateRange) {
        if (!"worksheet".equals(candidateKind) && !"table".equals(candidateKind)) {
            throw ServiceException.validation("AutoFilter owner kind is invalid");
        }
        if (candidateRange == null || !sheetId.equals(candidateRange.sheetId())) {
            throw ServiceException.validation("AutoFilter owner range must target its worksheet");
        }
        if ("table".equals(candidateKind) && (candidateTableId == null || candidateTableId.isBlank())) {
            throw ServiceException.validation("Table AutoFilter owner identity is required");
        }
        for (Owner owner : resolveOwners(sheet, sheetId)) {
            if (!owner.isSameOwner(candidateKind, candidateTableId) && overlaps(owner.range(), candidateRange)) {
                throw ServiceException.validation("AutoFilter owner cannot overlap " + ownerName(owner));
            }
        }
    }

    private static RangeRef rangeOf(JsonNode value, String sheetId, String label) {
        if (value == null || !value.isObject() || !sheetId.equals(value.path("sheetId").asText())) {
            throw ServiceException.validation(label + " range is invalid");
        }
        for (String coordinate : List.of("startRow", "endRow", "startColumn", "endColumn")) {
            if (!value.has(coordinate) || !value.get(coordinate).isIntegralNumber()) {
                throw ServiceException.validation(label + " range coordinates are required");
            }
        }
        try {
            return new RangeRef(sheetId, value.path("startRow").asInt(), value.path("endRow").asInt(), value.path("startColumn").asInt(), value.path("endColumn").asInt());
        } catch (IllegalArgumentException error) {
            throw ServiceException.validation(label + " range is invalid");
        }
    }

    private static String ownerName(Owner owner) {
        return owner.kind().equals("table") ? "Table AutoFilter " + owner.tableId() : "Worksheet AutoFilter";
    }

    private static boolean sameRange(RangeRef left, RangeRef right) {
        return left.sheetId().equals(right.sheetId())
                && left.startRow() == right.startRow()
                && left.endRow() == right.endRow()
                && left.startColumn() == right.startColumn()
                && left.endColumn() == right.endColumn();
    }

    private static boolean overlaps(RangeRef left, RangeRef right) {
        return left.startRow() <= right.endRow() && right.startRow() <= left.endRow()
                && left.startColumn() <= right.endColumn() && right.startColumn() <= left.endColumn();
    }
}
