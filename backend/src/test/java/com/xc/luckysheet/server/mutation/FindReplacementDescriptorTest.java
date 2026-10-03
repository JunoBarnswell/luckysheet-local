package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class FindReplacementDescriptorTest {
    private final ObjectMapper mapper = new ObjectMapper();
    private final FindReplacementDescriptor descriptor = new FindReplacementDescriptor();
    private ObjectNode snapshot() throws Exception {
        return (ObjectNode) mapper.readTree("{\"sheets\":[{\"id\":\"s\",\"rowCount\":10,\"columnCount\":10,\"cells\":{\"0\":{\"0\":{\"value\":\"before\"},\"1\":{\"value\":\"before\"}}}}]}");
    }
    private ObjectNode params() throws Exception {
        return (ObjectNode) mapper.readTree("{\"direction\":\"forward\",\"patches\":[{\"kind\":\"cell\",\"match\":{\"sheetId\":\"s\",\"row\":0,\"column\":0,\"key\":\"s:0:0\"},\"previous\":{\"value\":\"before\"},\"next\":{\"value\":\"=1+2\"}}]}");
    }
    @Test void actualBrowserMatchShapeKeepsReplacementInertAndSupportsUndo() throws Exception {
        var original = snapshot();
        var params = params();
        var operation = new OperationMutation("find.replaced", "s", params);
        assertEquals(1, descriptor.affectedRanges(original, operation).size());
        var result = descriptor.apply(original, operation);
        var cell = result.path("sheets").get(0).path("cells").path("0").path("0");
        assertEquals("=1+2", cell.path("value").asText());
        assertFalse(cell.has("formula"));
        assertEquals("before", original.path("sheets").get(0).path("cells").path("0").path("0").path("value").asText());
        params.put("direction", "reverse");
        assertEquals(original, descriptor.apply(result, new OperationMutation("find.replaced", "s", params)));
    }
    @Test void staleDuplicateAndMalformedMatchesRejectWithoutChangingSource() throws Exception {
        var original = snapshot();
        var preserved = original.deepCopy();
        var params = params();
        var patches = (com.fasterxml.jackson.databind.node.ArrayNode) params.get("patches");
        var second = ((ObjectNode) patches.get(0)).deepCopy();
        ((ObjectNode) second.get("match")).put("column", 1);
        ((ObjectNode) second.get("previous")).put("value", "stale");
        patches.add(second);
        assertThrows(ServiceException.class, () -> descriptor.apply(original, new OperationMutation("find.replaced", "s", params)));
        assertEquals(preserved, original);
        ((ObjectNode) second.get("match")).put("column", 0);
        assertThrows(ServiceException.class, () -> descriptor.affectedRanges(original, new OperationMutation("find.replaced", "s", params)));
        var malformed = params();
        ((ObjectNode) malformed.path("patches").get(0)).putNull("match");
        assertThrows(ServiceException.class, () -> descriptor.apply(original, new OperationMutation("find.replaced", "s", malformed)));
        assertEquals(preserved, original);
    }
}
