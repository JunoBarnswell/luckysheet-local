package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import java.time.Duration;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTimeout;

class WorkbookSnapshotValidatorTest {
    private final ObjectMapper mapper = new ObjectMapper();

    @Test
    void repeatedDrawingsDoNotRepeatedlyCanonicalizeAnEmptyConnectionSource() {
        ObjectNode snapshot = snapshot();
        ObjectNode sheet = (ObjectNode) snapshot.path("sheets").get(0);
        ObjectNode source = mapper.createObjectNode().put("kind", "table").put("tableId", "table-1");
        source.put("unrecognizedExtension", "x".repeat(500_000));
        sheet.withArray("pivots").add(pivot("pivot-1", source));
        sheet.withObject("drawingPayloads").set("payload-1", slicerPayload("pivot-1", mapper.createArrayNode()));
        ArrayNode drawings = sheet.withArray("drawings");
        for (int index = 0; index < 2_000; index++) {
            drawings.addObject().put("id", "drawing-" + index).put("payloadId", "payload-1");
        }

        assertTimeout(Duration.ofSeconds(2), () ->
                assertEquals(snapshot, WorkbookSnapshotValidator.requireCanonical(snapshot, "book-1")));
    }

    @Test
    void rejectsAConnectionWhoseSourceIdentityDoesNotMatch() {
        ObjectNode snapshot = snapshot();
        ObjectNode sheet = (ObjectNode) snapshot.path("sheets").get(0);
        ObjectNode source = mapper.createObjectNode().put("kind", "table").put("tableId", "table-1");
        sheet.withArray("pivots").add(pivot("pivot-1", source));
        sheet.withArray("pivots").add(pivot("pivot-2", source));
        ArrayNode connections = mapper.createArrayNode();
        connections.addObject().put("pivotId", "pivot-2").put("sourceKey", "attacker-controlled")
                .put("fieldId", "field-1");
        sheet.withObject("drawingPayloads").set("payload-1", slicerPayload("pivot-1", connections));
        sheet.withArray("drawings").addObject().put("id", "drawing-1").put("payloadId", "payload-1");

        ServiceException error = assertThrows(ServiceException.class,
                () -> WorkbookSnapshotValidator.requireCanonical(snapshot, "book-1"));
        assertEquals("VALIDATION_ERROR", error.code());
    }

    private ObjectNode snapshot() {
        ObjectNode snapshot = mapper.createObjectNode();
        snapshot.put("schema", GeneratedWorkbookContract.SNAPSHOT_SCHEMA).put("version", GeneratedWorkbookContract.SNAPSHOT_VERSION)
                .put("unitId", "book-1").put("name", "Book");
        snapshot.putObject("dimensionMetrics").put("normalFontFamily", "Calibri").put("normalFontSizePx", 14).put("maximumDigitWidthPx", 7);
        ObjectNode dataModel = snapshot.putObject("dataModel");
        dataModel.putArray("sources"); dataModel.putArray("tables"); dataModel.putArray("relationships"); dataModel.putArray("views");
        ObjectNode sheet = snapshot.putArray("sheets").addObject();
        sheet.put("kind", "worksheet").put("id", "sheet-1").put("name", "Sheet1").put("rowCount", 10).put("columnCount", 10)
                .put("defaultRowHeightPx", 20).put("defaultColumnWidthPx", 64);
        sheet.putObject("cells"); sheet.putArray("merges"); sheet.putObject("pane").put("kind", "none");
        sheet.putArray("pivots"); sheet.putArray("sparklines"); sheet.putArray("drawings"); sheet.putObject("drawingPayloads");
        return snapshot;
    }

    private ObjectNode pivot(String id, ObjectNode source) {
        ObjectNode pivot = mapper.createObjectNode().put("id", id);
        pivot.set("source", source.deepCopy());
        pivot.putObject("fieldCatalog").putArray("fields").addObject()
                .put("fieldId", "field-1").put("ordinal", 0).put("name", "Region").put("dataType", "text");
        return pivot;
    }

    private ObjectNode slicerPayload(String pivotId, ArrayNode connections) {
        ObjectNode payload = mapper.createObjectNode().put("kind", "slicer").put("pivotId", pivotId).put("fieldId", "field-1");
        payload.set("connections", connections);
        return payload;
    }
}
