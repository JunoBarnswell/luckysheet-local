package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.store.WorkbookDataBlockStore;
import org.junit.jupiter.api.Test;

import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verifyNoInteractions;
import static org.mockito.Mockito.when;

class WorkbookDataBlockPublicationGuardTest {
    private final ObjectMapper mapper = new ObjectMapper();

    @Test
    void capturedBlockReferencesSurviveOwnedSnapshotMutation() throws Exception {
        WorkbookDataBlockStore store = mock(WorkbookDataBlockStore.class);
        WorkbookDataBlockPublicationGuard guard = new WorkbookDataBlockPublicationGuard(store);
        ObjectNode previous = snapshot("checksum-a");
        JsonNode candidate = previous.deepCopy();
        WorkbookDataBlockPublicationGuard.PreviousBlockReferences references = guard.capturePreviousBlockReferences(previous);
        ((ObjectNode) previous.path("dataModel").path("sources").get(0).path("blocks").get(0)).put("checksum", "mutated-in-place");

        guard.requireNewReferences("unit-1", candidate, references);

        verifyNoInteractions(store);
    }

    @Test
    void changedCapturedBlockReferenceStillFailsClosed() throws Exception {
        WorkbookDataBlockStore store = mock(WorkbookDataBlockStore.class);
        WorkbookDataBlockPublicationGuard guard = new WorkbookDataBlockPublicationGuard(store);
        ObjectNode previous = snapshot("checksum-a");
        JsonNode candidate = previous.deepCopy();
        WorkbookDataBlockPublicationGuard.PreviousBlockReferences references = guard.capturePreviousBlockReferences(previous);
        ((ObjectNode) candidate.path("dataModel").path("sources").get(0).path("blocks").get(0)).put("checksum", "checksum-b");
        when(store.findMetadata("unit-1", "source-1", "block-1")).thenReturn(Optional.empty());

        ServiceException error = assertThrows(ServiceException.class,
                () -> guard.requireNewReferences("unit-1", candidate, references));

        assertEquals("DATA_BLOCK_MISSING", error.code());
    }

    private ObjectNode snapshot(String checksum) throws Exception {
        return (ObjectNode) mapper.readTree("""
                {"dataModel":{"sources":[{"id":"source-1","blocks":[{"id":"block-1","dataSourceId":"source-1","byteLength":10,"checksum":"%s"}]}]}}
                """.formatted(checksum));
    }
}
