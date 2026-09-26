package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ProtectionResolverOwnedPreimageTest {
    private final ObjectMapper mapper = new ObjectMapper();

    @Test
    void protectionPreimagePreservesPermissionDecisionsWithoutCopyingCellPayload() throws Exception {
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":20,"columnCount":10,
                  "cells":{"0":{"0":{"value":"large cell payload"}}},
                  "protectionRules":[{"id":"locked-range","scope":"range","sheetId":"sheet-1",
                    "range":{"sheetId":"sheet-1","startRow":5,"endRow":7,"startColumn":0,"endColumn":3},
                    "locked":true,"allow":{"insertRows":false}}]}]}
                """);
        JsonNode preimage = ProtectionResolver.structuralProtectionPreimage(snapshot);

        assertTrue(preimage.path("sheets").get(0).path("cells").isMissingNode());
        assertDoesNotThrow(() -> ProtectionResolver.assertAllowed(preimage,
                List.of(new RangeRef("sheet-1", 10, 10, 0, 3)), "insert-rows"));
        ServiceException error = assertThrows(ServiceException.class, () -> ProtectionResolver.assertAllowed(preimage,
                List.of(new RangeRef("sheet-1", 6, 6, 1, 1)), "insert-rows"));
        assertEquals("FORBIDDEN", error.code());
    }

    @Test
    void rejectsMissingProtectionActionEvenWhenThereAreNoAffectedRanges() throws Exception {
        JsonNode snapshot = mapper.readTree("{\"sheets\":[]}");

        ServiceException error = assertThrows(ServiceException.class,
                () -> ProtectionResolver.assertAllowed(snapshot, List.of(), null));

        assertEquals("VALIDATION_ERROR", error.code());
    }
}
