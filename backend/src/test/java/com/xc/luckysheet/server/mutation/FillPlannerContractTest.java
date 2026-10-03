package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class FillPlannerContractTest {
    @Test void wholeAxisCopyPreservesAbsoluteEndpointsAndQuotedText() {
        assertEquals("=SUM(2:2)+SUM(B:B)+SUM($A:C)+\"1:1\"", FormulaReferenceTransformer.offsetForCopy("=SUM(1:1)+SUM(A:A)+SUM($A:B)+\"1:1\"", 1, 1));
        assertThrows(ServiceException.class, () -> FormulaReferenceTransformer.offsetForCopy("=SUM(1:1)", -1, 0));
        assertThrows(ServiceException.class, () -> FormulaReferenceTransformer.offsetForCopy("='[External.xlsx]Sheet'!A1", 1, 0));
    }
    @Test void serverMatchesIndependentTypeScriptPlansAndRejectsForgedResults() throws Exception {
        var mapper = new ObjectMapper();
        try (var input = getClass().getResourceAsStream("/security/fill-plans.json")) {
            assertNotNull(input);
            var cases = mapper.readTree(input);
            assertEquals(28, cases.size());
            for (JsonNode fixture : cases) {
                var root = fixture.path("snapshot"); var original = root.deepCopy();
                var params = (ObjectNode) fixture.path("params").deepCopy();
                String sheetId = params.path("sheetId").asText();
                var reducer = new FillMutationDescriptor("fill.applied");
                var result = reducer.apply(root, new OperationMutation("fill.applied", sheetId, params));
                for (var write : params.path("writes")) {
                    var cell = result.path("sheets").get(0).path("cells").path(write.path("row").asText()).path(write.path("column").asText());
                    var expected = write.path("after");
                    if (expected.path("value").isNumber()) assertEquals(expected.path("value").asDouble(), cell.path("value").asDouble(), fixture.path("name").asText());
                    else assertEquals(expected.path("value"), cell.path("value"), fixture.path("name").asText());
                    assertEquals(expected.path("numberFormat"), cell.path("numberFormat"));
                }
                assertEquals(original, root);
                ((ObjectNode) params.path("writes").get(0).path("after")).put("value", "forged-result");
                assertThrows(ServiceException.class, () -> reducer.apply(root, new OperationMutation("fill.applied", sheetId, params)), fixture.path("name").asText());
                assertEquals(original, root);
            }
        }
    }
}
