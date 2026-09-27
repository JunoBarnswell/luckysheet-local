package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class DataRegionContextValidatorTest {
    private static final String SHEET_ID = "sheet-1";
    private final ObjectMapper mapper = new ObjectMapper();

    @Test
    void worksheetSortCanOverrideTheInferredHeaderInEitherDirection() {
        ObjectNode inferredPresent = sortParams("worksheet", null, "present");
        inferredPresent.put("hasHeader", false);
        assertDoesNotThrow(() -> DataRegionContextValidator.validateSort(sheetRoot(false), SHEET_ID, inferredPresent));

        ObjectNode inferredAbsent = sortParams("worksheet", null, "absent");
        inferredAbsent.put("hasHeader", true);
        assertDoesNotThrow(() -> DataRegionContextValidator.validateSort(sheetRoot(false), SHEET_ID, inferredAbsent));
    }

    @Test
    void sheetTableSortCannotOverrideItsCanonicalHeaderMetadata() {
        ObjectNode params = sortParams("sheet-table", "table-1", "present");
        params.put("hasHeader", false);

        ServiceException error = assertThrows(ServiceException.class,
                () -> DataRegionContextValidator.validateSort(sheetRoot(true), SHEET_ID, params));
        assertEquals("VALIDATION_ERROR", error.code());
    }

    private ObjectNode sheetRoot(boolean withTable) {
        ObjectNode root = mapper.createObjectNode();
        ObjectNode sheet = root.putArray("sheets").addObject();
        sheet.put("id", SHEET_ID);
        if (withTable) {
            ObjectNode table = sheet.putArray("sheetTables").addObject();
            table.put("id", "table-1").put("hasHeaderRow", true);
            table.set("range", range());
        }
        return root;
    }

    private ObjectNode sortParams(String ownerKind, String tableId, String headerKind) {
        ObjectNode params = mapper.createObjectNode();
        params.set("range", range());
        ObjectNode context = params.putObject("dataRegionContext");
        context.put("schema", "DataRegionContext").put("version", 1);
        context.set("selection", range());
        context.set("currentRegion", range());
        context.set("range", range());
        context.set("usedRange", range());
        context.put("activeColumn", 0).put("searchScope", "current-region");
        ObjectNode owner = context.putObject("owner").put("kind", ownerKind);
        if (tableId != null) owner.put("tableId", tableId);
        ObjectNode header = context.putObject("header").put("kind", headerKind);
        if ("present".equals(headerKind)) header.put("row", 0);
        context.putArray("visibleRows");
        return params;
    }

    private ObjectNode range() {
        return mapper.createObjectNode().put("sheetId", SHEET_ID)
                .put("startRow", 0).put("endRow", 2).put("startColumn", 0).put("endColumn", 1);
    }
}
