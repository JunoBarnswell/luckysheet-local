package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RecordTableValidator;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class RecordMutationDescriptorTest {
    private final ObjectMapper mapper = new ObjectMapper();
    private JsonNode fixture() throws Exception {
        return mapper.readTree("""
          {"sheets":[{"id":"source","name":"Source","rowCount":100,"columnCount":10,"cells":{
            "1":{"0":{"value":"o1"},"1":{"value":"c1"},"2":{"value":2}},
            "2":{"0":{"value":"o2"},"1":{"value":"c1"},"2":{"value":3}},
            "5":{"0":{"value":"c1"},"1":{"value":"Alpha"}}}}],
          "dataModel":{"tables":[
            {"id":"orders","name":"Orders","recordIdFieldId":"id","sourceRange":{"sheetId":"source","startRow":0,"endRow":2,"startColumn":0,"endColumn":3},
             "fields":[{"id":"id","name":"ID","ordinal":0,"type":"text"},{"id":"customer","name":"Customer","ordinal":1,"type":"text"},{"id":"quantity","name":"Quantity","ordinal":2,"type":"number"},{"id":"amount","name":"Amount","ordinal":3,"type":"number","calculation":{"kind":"formula","formula":"=[@quantity]*10"}}]},
            {"id":"customers","name":"Customers","recordIdFieldId":"id","sourceRange":{"sheetId":"source","startRow":4,"endRow":5,"startColumn":0,"endColumn":2},
             "fields":[{"id":"id","name":"ID","ordinal":0,"type":"text"},{"id":"name","name":"Name","ordinal":1,"type":"text"},{"id":"total","name":"Total","ordinal":2,"type":"number","calculation":{"kind":"rollup","relationshipId":"relation","targetFieldId":"amount","direction":"reverse","aggregate":"SUM"}}]}],
             "relationships":[{"id":"relation","fromTableId":"orders","fromFieldId":"customer","toTableId":"customers","toFieldId":"id","cardinality":"many-to-one"}]}}
          """);
    }
    private OperationMutation write(String record, String field, int row, int column, JsonNode value) {
        ObjectNode cell = mapper.createObjectNode(); cell.set("value", value);
        ObjectNode target = mapper.createObjectNode().put("sheetId", "source").put("row", row).put("column", column);
        ObjectNode authority = mapper.createObjectNode().put("kind", "direct-entry");
        authority.set("target", target); authority.set("candidate", cell); authority.putObject("validationDecision").put("status", "accepted");
        ObjectNode params = mapper.createObjectNode().put("tableId", "orders").put("recordId", record).put("fieldId", field);
        params.set("value", value); params.set("writeAuthority", authority);
        return new OperationMutation("record.set", "source", params);
    }
    @Test void recordWritesResolveStableIdsAndRejectComputedAndMissingRelationTargets() throws Exception {
        JsonNode before = fixture(); RecordTableValidator.validateWorkbook(before);
        var descriptor = new RecordMutationDescriptor("record.set");
        JsonNode next = descriptor.apply(before, write("o2", "quantity", 2, 2, mapper.valueToTree(7)));
        assertEquals(7, next.path("sheets").get(0).path("cells").path("2").path("2").path("value").asInt());
        assertEquals(3, before.path("sheets").get(0).path("cells").path("2").path("2").path("value").asInt());
        assertThrows(ServiceException.class, () -> descriptor.apply(before, write("o1", "amount", 1, 3, mapper.valueToTree(9))));
        assertThrows(ServiceException.class, () -> descriptor.apply(before, write("o1", "customer", 1, 1, mapper.valueToTree("missing"))));
        assertEquals("c1", before.path("sheets").get(0).path("cells").path("1").path("1").path("value").asText());
    }
    @Test void duplicateIdentityMissingLookupFieldAndPhysicalWritesFailClosed() throws Exception {
        JsonNode before = fixture();
        ObjectNode duplicate = (ObjectNode) before.deepCopy();
        ((ObjectNode) duplicate.path("sheets").get(0).path("cells").path("2").path("0")).put("value", "o1");
        assertThrows(ServiceException.class, () -> RecordTableValidator.validateWorkbook(duplicate));
        ObjectNode invalid = (ObjectNode) before.deepCopy();
        ((ObjectNode) invalid.path("dataModel").path("tables").get(1).path("fields").get(2).path("calculation")).put("targetFieldId", "missing");
        assertThrows(ServiceException.class, () -> RecordTableValidator.validateWorkbook(invalid));
        OperationMutation physical = new OperationMutation("range.clear", "source", mapper.readTree("""
          {"range":{"sheetId":"source","startRow":1,"endRow":1,"startColumn":0,"endColumn":0}}
          """));
        assertThrows(ServiceException.class, () -> new MutationDescriptorRegistry().prepare(before, physical, WorkbookAclRole.OWNER));
    }
}
