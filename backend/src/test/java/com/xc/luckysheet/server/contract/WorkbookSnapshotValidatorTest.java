package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import java.time.Duration;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
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
            drawings.addObject().put("id", "drawing-" + index).put("sheetId", "sheet-1").put("kind", "slicer").put("payloadId", "payload-1")
                    .putObject("anchor").put("kind", "absolute");
            ((ObjectNode) drawings.get(index)).putObject("transform").put("x", 0).put("y", 0).put("width", 100).put("height", 40).put("rotation", 0);
            ((ObjectNode) drawings.get(index)).put("zIndex", index);
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
        sheet.withArray("drawings").addObject().put("id", "drawing-1").put("sheetId", "sheet-1").put("kind", "slicer").put("payloadId", "payload-1")
                .putObject("anchor").put("kind", "absolute");
        ((ObjectNode) sheet.path("drawings").get(0)).putObject("transform").put("x", 0).put("y", 0).put("width", 100).put("height", 40).put("rotation", 0);
        ((ObjectNode) sheet.path("drawings").get(0)).put("zIndex", 0);

        ServiceException error = assertThrows(ServiceException.class,
                () -> WorkbookSnapshotValidator.requireCanonical(snapshot, "book-1"));
        assertEquals("VALIDATION_ERROR", error.code());
    }

    @Test
    void migratesV9DefinedNamesAndCellHyperlinksIntoCanonicalV10Owners() {
        ObjectNode legacy = snapshot().put("version", 9);
        legacy.putObject("definedNames").put("SalesTotal", "=Sheet1!A1");
        ObjectNode cell = ((ObjectNode) legacy.path("sheets").get(0)).withObject("cells").putObject("0").putObject("0");
        cell.put("value", "Sales").put("hyperlink", "https://example.test/sales");

        ObjectNode migrated = WorkbookSnapshotValidator.migrateStored(legacy, "book-1");

        assertEquals(10, migrated.path("version").asInt());
        assertFalse(migrated.has("definedNames"));
        assertEquals("SalesTotal", migrated.path("definedNameModels").get(0).path("name").asText());
        assertEquals("https://example.test/sales", migrated.path("sheets").get(0).path("hyperlinks").get(0).path("hyperlink").path("target").path("url").asText());
        assertFalse(migrated.path("sheets").get(0).path("cells").path("0").path("0").has("hyperlink"));
    }

    @Test
    void rejectsV9ConflictingDefinedNameAndHyperlinkRepresentations() {
        ObjectNode legacy = snapshot().put("version", 9);
        legacy.putObject("definedNames").put("SalesTotal", "=Sheet1!A1");
        legacy.withArray("definedNameModels").addObject().put("name", "SalesTotal").put("formula", "=Sheet1!B1").put("scope", "workbook");
        ObjectNode cell = ((ObjectNode) legacy.path("sheets").get(0)).withObject("cells").putObject("0").putObject("0");
        cell.put("value", "Sales").put("hyperlink", "https://example.test/a");
        cell.putObject("hyperlinkDetail").put("id", "link-1").putObject("target").put("kind", "url").put("url", "https://example.test/b");

        ServiceException error = assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.migrateStored(legacy, "book-1"));
        assertEquals("VALIDATION_ERROR", error.code());
    }

    @Test
    void requireCanonicalRejectsRemovedV9Fields() {
        ObjectNode legacyDefinedNames = snapshot();
        legacyDefinedNames.putObject("definedNames");
        assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.requireCanonical(legacyDefinedNames, "book-1"));

        ObjectNode legacyCell = snapshot();
        ((ObjectNode) legacyCell.path("sheets").get(0)).withObject("cells").putObject("0").putObject("0").put("hyperlink", "https://example.test");
        assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.requireCanonical(legacyCell, "book-1"));
    }

    private ObjectNode snapshot() {
        ObjectNode snapshot = mapper.createObjectNode();
        snapshot.put("schema", GeneratedWorkbookContract.SNAPSHOT_SCHEMA).put("version", GeneratedWorkbookContract.SNAPSHOT_VERSION)
                .put("unitId", "book-1").put("name", "Book");
        snapshot.putObject("dimensionMetrics").put("normalFontFamily", "Calibri").put("normalFontSizePx", 14).put("maximumDigitWidthPx", 7);
        snapshot.putObject("calculationSettings").put("mode", "automatic").put("iterativeCalculation", false).put("maximumIterations", 100)
                .put("maximumChange", 0.001).put("precisionAsDisplayed", false).put("calculateBeforeSave", true).put("fullCalculationOnLoad", false);
        snapshot.putObject("editingOptions").put("allowEditDirectly", true).put("moveAfterEnter", true).put("enterDirection", "down")
                .put("formulaAutoComplete", true).put("valueAutoComplete", true).putNull("fixedDecimalPlaces");
        ObjectNode dataModel = snapshot.putObject("dataModel");
        dataModel.putArray("sources"); dataModel.putArray("tables"); dataModel.putArray("relationships"); dataModel.putArray("views");
        snapshot.putArray("definedNameModels");
        ObjectNode sheet = snapshot.putArray("sheets").addObject();
        sheet.put("kind", "worksheet").put("id", "sheet-1").put("name", "Sheet1").put("rowCount", 10).put("columnCount", 10)
                .put("defaultRowHeightPx", 20).put("defaultColumnWidthPx", 64);
        sheet.putObject("cells"); sheet.putArray("merges"); sheet.putObject("pane").put("kind", "none");
        sheet.putArray("pivots"); sheet.putArray("sparklines"); sheet.putArray("drawings"); sheet.putObject("drawingPayloads");
        ObjectNode review = sheet.putObject("review");
        review.putObject("notesByCell"); review.putObject("notesById"); review.putObject("threadIdsByCell"); review.putObject("threadsById");
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
