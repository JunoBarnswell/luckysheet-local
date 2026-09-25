package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.migration.SnapshotUpgrade;
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

    @Test
    void rejectsCellsOutsideTheDeclaredWorksheetExtent() {
        ObjectNode snapshot = snapshot();
        ((ObjectNode) snapshot.path("sheets").get(0)).withObject("cells").withObject("10").putObject("0").put("value", "outside");

        ServiceException error = assertThrows(ServiceException.class,
                () -> WorkbookSnapshotValidator.requireCanonical(snapshot, "book-1"));

        assertEquals("VALIDATION_ERROR", error.code());
    }

    @Test
    void rejectsPaneCoordinatesAndFrozenSplitsOutsideTheirCanonicalDomains() throws Exception {
        for (String pane : java.util.List.of(
                "{\"kind\":\"none\",\"activePane\":\"center\"}",
                "{\"kind\":\"frozen\",\"state\":\"frozen\",\"xSplit\":1,\"ySplit\":0,\"startRow\":0}",
                "{\"kind\":\"frozen\",\"state\":\"frozen\",\"xSplit\":1.5,\"ySplit\":0,\"startRow\":0,\"startColumn\":1}",
                "{\"kind\":\"frozen\",\"state\":\"frozen\",\"xSplit\":1,\"ySplit\":0,\"startRow\":0,\"startColumn\":16384}",
                "{\"kind\":\"split\",\"state\":\"split\",\"xSplit\":-0.5,\"ySplit\":10,\"startRow\":0,\"startColumn\":0}",
                "{\"kind\":\"split\",\"state\":\"split\",\"xSplit\":20,\"ySplit\":10,\"startRow\":1048576,\"startColumn\":0}",
                "{\"kind\":\"split\",\"state\":\"split\",\"xSplit\":20,\"ySplit\":10,\"startRow\":0,\"startColumn\":0,\"activePane\":\"center\"}",
                "{\"kind\":\"frozen\",\"state\":\"frozen\",\"xSplit\":1,\"ySplit\":0,\"startRow\":0,\"startColumn\":1,\"referenceHint\":\"A1\"}",
                "{\"kind\":\"none\",\"referenceHint\":\"A1\"}")) {
            ObjectNode candidate = snapshot();
            ((ObjectNode) candidate.path("sheets").get(0)).set("pane", mapper.readTree(pane));

            ServiceException error = assertThrows(ServiceException.class,
                    () -> WorkbookSnapshotValidator.requireCanonical(candidate, "book-1"));
            assertEquals("VALIDATION_ERROR", error.code());
        }
    }

    @Test
    void migratesV9CellHyperlinksToTheCanonicalWorksheetOwner() {
        ObjectNode snapshot = snapshot().put("version", 9);
        ObjectNode sheet = (ObjectNode) snapshot.path("sheets").get(0);
        ObjectNode row = sheet.withObject("cells").putObject("0");
        row.putObject("0").put("value", "old").put("hyperlink", "https://legacy.example");
        row.putObject("1").put("value", "detail").set("hyperlinkDetail", mapper.createObjectNode()
                .put("id", "legacy-detail").set("target", mapper.createObjectNode().put("kind", "email").put("address", "link@example.com")));
        sheet.withArray("hyperlinks").addObject().put("row", 0).put("column", 0)
                .set("hyperlink", mapper.createObjectNode().put("id", "canonical")
                        .set("target", mapper.createObjectNode().put("kind", "url").put("url", "https://canonical.example")));

        ObjectNode migrated = SnapshotUpgrade.migrateStored(snapshot, "book-1");
        ObjectNode migratedSheet = (ObjectNode) migrated.path("sheets").get(0);

        assertEquals(GeneratedWorkbookContract.SNAPSHOT_VERSION, migrated.path("version").intValue());
        assertEquals(2, migratedSheet.path("hyperlinks").size());
        assertEquals("canonical", migratedSheet.path("hyperlinks").get(1).path("hyperlink").path("id").asText());
        assertEquals("legacy-detail", migratedSheet.path("hyperlinks").get(0).path("hyperlink").path("id").asText());
        assertEquals(false, migratedSheet.path("cells").path("0").path("0").has("hyperlink"));
        assertEquals(false, migratedSheet.path("cells").path("0").path("1").has("hyperlinkDetail"));
    }

    @Test
    void rejectsLegacyCellHyperlinkMetadataInCanonicalSnapshots() {
        ObjectNode snapshot = snapshot();
        ((ObjectNode) snapshot.path("sheets").get(0)).withObject("cells").putObject("0").putObject("0")
                .put("hyperlink", "https://legacy.example");

        ServiceException error = assertThrows(ServiceException.class,
                () -> WorkbookSnapshotValidator.requireCanonical(snapshot, "book-1"));

        assertEquals("VALIDATION_ERROR", error.code());
    }

    @Test
    void requiresSheetTableColumnCountToMatchItsRangeWidth() {
        ObjectNode snapshot = snapshot();
        ObjectNode sheet = (ObjectNode) snapshot.path("sheets").get(0);
        ObjectNode table = sheet.putArray("sheetTables").addObject();
        table.put("id", "table-1").put("sheetId", "sheet-1").put("name", "Table1");
        table.putObject("range").put("sheetId", "sheet-1").put("startRow", 0).put("endRow", 2)
                .put("startColumn", 1).put("endColumn", 2);
        ArrayNode columns = table.putArray("columns");
        columns.addObject().put("id", "column-1").put("name", "A");
        columns.addObject().put("id", "column-2").put("name", "B");
        assertEquals(snapshot, WorkbookSnapshotValidator.requireCanonical(snapshot, "book-1"));
        columns.remove(1);

        ServiceException error = assertThrows(ServiceException.class,
                () -> WorkbookSnapshotValidator.requireCanonical(snapshot, "book-1"));

        assertEquals("VALIDATION_ERROR", error.code());
    }

    @Test
    void rejectsDanglingWorksheetHyperlinkTargetsAndDuplicateAnchors() {
        ObjectNode snapshot = snapshot();
        ObjectNode sheet = (ObjectNode) snapshot.path("sheets").get(0);
        ObjectNode link = mapper.createObjectNode().put("id", "dangling");
        link.set("target", mapper.createObjectNode().put("kind", "sheet").put("sheetId", "missing-sheet").put("address", "A1"));
        sheet.withArray("hyperlinks").addObject().put("row", 0).put("column", 0).set("hyperlink", link);

        ServiceException dangling = assertThrows(ServiceException.class,
                () -> WorkbookSnapshotValidator.requireCanonical(snapshot, "book-1"));
        assertEquals("VALIDATION_ERROR", dangling.code());

        sheet.withArray("hyperlinks").removeAll();
        sheet.withArray("hyperlinks").addObject().put("row", 0).put("column", 0)
                .set("hyperlink", mapper.createObjectNode().put("id", "duplicate")
                        .set("target", mapper.createObjectNode().put("kind", "url").put("url", "https://example.com")));
        sheet.withArray("hyperlinks").addObject().put("row", 0).put("column", 0)
                .set("hyperlink", mapper.createObjectNode().put("id", "duplicate-2")
                        .set("target", mapper.createObjectNode().put("kind", "url").put("url", "https://example.org")));
        ServiceException duplicate = assertThrows(ServiceException.class,
                () -> WorkbookSnapshotValidator.requireCanonical(snapshot, "book-1"));
       assertEquals("VALIDATION_ERROR", duplicate.code());
   }

    @Test
    void rejectsUndeclaredHyperlinkTargetFields() {
        ObjectNode snapshot = snapshot();
        ObjectNode target = mapper.createObjectNode().put("kind", "url").put("url", "https://example.com").put("sheetId", "unexpected");
        ObjectNode hyperlink = mapper.createObjectNode().put("id", "link");
        hyperlink.set("target", target);
        ((ObjectNode) snapshot.path("sheets").get(0)).withArray("hyperlinks").addObject()
                .put("row", 0).put("column", 0).set("hyperlink", hyperlink);

        ServiceException error = assertThrows(ServiceException.class,
                () -> WorkbookSnapshotValidator.requireCanonical(snapshot, "book-1"));

        assertEquals("VALIDATION_ERROR", error.code());
    }

    @Test
    void rejectsNonTextualIdsAndNonCanonicalTargetKinds() {
        ObjectNode numericSnapshot = snapshot();
        ObjectNode numericId = mapper.createObjectNode().put("id", 7);
        numericId.set("target", mapper.createObjectNode().put("kind", "url").put("url", "https://example.com"));
        ((ObjectNode) numericSnapshot.path("sheets").get(0)).withArray("hyperlinks").addObject()
                .put("row", 0).put("column", 0).set("hyperlink", numericId);
        ServiceException invalidId = assertThrows(ServiceException.class,
                () -> WorkbookSnapshotValidator.requireCanonical(numericSnapshot, "book-1"));
        assertEquals("VALIDATION_ERROR", invalidId.code());

        ObjectNode kindSnapshot = snapshot();
        ObjectNode nonCanonicalKind = mapper.createObjectNode().put("id", "link");
        nonCanonicalKind.set("target", mapper.createObjectNode().put("kind", " url ").put("url", "https://example.com"));
        ((ObjectNode) kindSnapshot.path("sheets").get(0)).withArray("hyperlinks").addObject()
                .put("row", 0).put("column", 0).set("hyperlink", nonCanonicalKind);
        ServiceException invalidKind = assertThrows(ServiceException.class,
                () -> WorkbookSnapshotValidator.requireCanonical(kindSnapshot, "book-1"));
        assertEquals("VALIDATION_ERROR", invalidKind.code());
    }

    @Test
    void canonicalDefinedNameModelsRemainTheOnlyAuthorityForNameHyperlinks() {
        ObjectNode snapshot = snapshot();
        snapshot.putObject("definedNames").put("StaleProjection", "=A1");
        snapshot.putArray("definedNameModels");
        ObjectNode target = mapper.createObjectNode().put("kind", "name").put("name", "StaleProjection");
        ObjectNode hyperlink = mapper.createObjectNode().put("id", "name-link");
        hyperlink.set("target", target);
        ((ObjectNode) snapshot.path("sheets").get(0)).withArray("hyperlinks").addObject()
                .put("row", 0).put("column", 0).set("hyperlink", hyperlink);

        ServiceException error = assertThrows(ServiceException.class,
                () -> WorkbookSnapshotValidator.requireCanonical(snapshot, "book-1"));

        assertEquals("VALIDATION_ERROR", error.code());
    }

    @Test
    void rejectsAutoFilterColumnKeysWithNumericAliases() {
        ObjectNode snapshot = snapshot();
        ObjectNode filter = ((ObjectNode) snapshot.path("sheets").get(0)).putObject("autoFilter");
        filter.put("sheetId", "sheet-1");
        filter.set("range", mapper.createObjectNode().put("sheetId", "sheet-1")
                .put("startRow", 0).put("endRow", 4).put("startColumn", 0).put("endColumn", 2));
        ObjectNode columns = filter.putObject("columns");
        columns.putObject("1").put("column", 1).put("showButton", true).put("hiddenButton", false);
        columns.putObject("01").put("column", 1).put("showButton", true).put("hiddenButton", false);

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
        sheet.putArray("pivots"); sheet.putArray("sparklines"); sheet.putArray("drawings"); sheet.putObject("drawingPayloads"); sheet.putArray("hyperlinks");
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
