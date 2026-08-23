package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class MutationDescriptorRegistryTest {
    private final ObjectMapper mapper = new ObjectMapper();

    @Test
    void unknownMutationsFailClosed() {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ServiceException error = assertThrows(ServiceException.class, () -> registry.require("unknown.mutation", false));
        assertEquals("VALIDATION_ERROR", error.code());
    }

    @Test
    void cellSetUsesServerResolvedRangeAndChangesSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("{\"sheets\":[{\"id\":\"sheet-1\",\"cells\":{}}]}");
        var mutation = new OperationMutation("cell.set", "sheet-1", mapper.readTree("{\"row\":2,\"column\":3,\"value\":{\"value\":42}}"));
        assertEquals(2, registry.resolveRanges(snapshot, mutation).get(0).startRow());
        var next = registry.applyPublicMutations(snapshot, List.of(mutation));
        assertEquals(42, next.path("sheets").get(0).path("cells").path("2").path("3").path("value").asInt());
    }

    @Test
    void internalRestoreCannotBeSubmittedByClient() {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        assertThrows(ServiceException.class, () -> registry.require("workbook.restore", false));
    }
}
