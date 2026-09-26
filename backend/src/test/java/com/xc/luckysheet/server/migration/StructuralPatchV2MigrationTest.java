package com.xc.luckysheet.server.migration;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.CommittedOperationMutation;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.OperationOrigin;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.StructuralPatch;
import com.xc.luckysheet.server.mutation.MutationDescriptorRegistry;
import org.junit.jupiter.api.Test;

import java.time.Instant;
import java.util.List;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class StructuralPatchV2MigrationTest {
    private final ObjectMapper mapper = new ObjectMapper();

    private static final class V4__TestMigration extends StructuralPatchV2Migration { }

    @Test
    void upgradesLegacyPatchVersionsWithoutComparingNewRangeImpactsToOldImpactLists() throws Exception {
        for (int version : List.of(1, 2, 3)) {
            MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
            StructuralPatch currentPatch = currentRangePatch();
            ObjectNode rawEnvelope = rawEnvelope(version, currentPatch, List.of());

            new V4__TestMigration().rewriteMutationPatches(rawEnvelope, operation(),
                    List.of(Optional.of(currentPatch)), registry, "unit-1", 1);

            ObjectNode migratedMutation = (ObjectNode) rawEnvelope.path("mutations").get(0);
            assertEquals(mapper.valueToTree(currentPatch), migratedMutation.get("structuralPatch"));
            assertEquals(mapper.valueToTree(registry.structuralImpactRanges(currentPatch)),
                    migratedMutation.get("structuralImpactRanges"));
        }
    }

    @Test
    void rejectsLegacyImpactTamperingAndUpgradesCanonicalV4RangeImpacts() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        StructuralPatch currentPatch = currentRangePatch();
        List<RangeRef> preV5Impact = preV5RangeImpacts(currentPatch, registry);

        ObjectNode tamperedLegacyEnvelope = rawEnvelope(3, currentPatch, registry.structuralImpactRanges(currentPatch));
        IllegalStateException mismatch = assertThrows(IllegalStateException.class,
                () -> new V4__TestMigration().rewriteMutationPatches(tamperedLegacyEnvelope, operation(),
                        List.of(Optional.of(currentPatch)), registry, "unit-1", 1));
        assertTrue(mismatch.getMessage().contains("STRUCTURAL_IMPACT_MISMATCH"));

        ObjectNode currentEnvelope = rawEnvelope(4, currentPatch, preV5Impact);
        new V4__TestMigration().rewriteMutationPatches(currentEnvelope, operation(),
                List.of(Optional.of(currentPatch)), registry, "unit-1", 1);
        assertEquals(mapper.valueToTree(currentPatch), currentEnvelope.path("mutations").get(0).get("structuralPatch"));

        ObjectNode canonicalV5Envelope = rawEnvelope(5, currentPatch, registry.structuralImpactRanges(currentPatch));
        new V4__TestMigration().rewriteMutationPatches(canonicalV5Envelope, operation(),
                List.of(Optional.of(currentPatch)), registry, "unit-1", 1);
        assertEquals(mapper.valueToTree(currentPatch), canonicalV5Envelope.path("mutations").get(0).get("structuralPatch"));
    }

    private ObjectNode rawEnvelope(int version, StructuralPatch currentPatch, List<RangeRef> impact) {
        ObjectNode legacyPatch = (ObjectNode) mapper.valueToTree(currentPatch);
        legacyPatch.put("version", version);
        if (version < 4) legacyPatch.remove("rangeOwnerDeltas");
        if (version == 1) legacyPatch.remove("definedNameOwnerDeltas");
        ObjectNode envelope = mapper.createObjectNode();
        ArrayNode mutations = envelope.putArray("mutations");
        ObjectNode mutation = mutations.addObject().put("id", "rows.inserted").put("sheetId", "sheet-1");
        if (version == 4) {
            ArrayNode preV5RangeOwners = mapper.createArrayNode();
            for (JsonNode delta : legacyPatch.path("rangeOwnerDeltas")) {
                if (!"sheet-table".equals(delta.path("ownerKind").asText())) preV5RangeOwners.add(delta.deepCopy());
            }
            legacyPatch.set("rangeOwnerDeltas", preV5RangeOwners);
        }
        mutation.set("structuralPatch", legacyPatch);
        mutation.set("structuralImpactRanges", mapper.valueToTree(impact));
        return envelope;
    }

    private CommittedOperationEnvelope operation() {
        Instant now = Instant.parse("2026-09-26T00:00:00Z");
        CommittedOperationMutation mutation = new CommittedOperationMutation(
                "rows.inserted", "sheet-1", mapper.createObjectNode(), List.of());
        return new CommittedOperationEnvelope("session-1", OperationEnvelope.SCHEMA, "operation-1", "unit-1",
                "actor-1", OperationOrigin.CLIENT, 1, 0, 1, List.of(mutation), now, now);
    }

    private StructuralPatch currentRangePatch() {
        StructuralPatch.RangeOwnerDelta tableDelta = StructuralPatch.RangeOwnerDelta.range(
                "workbook-table", "table-1",
                new RangeRef("sheet-1", 0, 2, 0, 1),
                new RangeRef("sheet-1", 1, 3, 0, 1));
        StructuralPatch.RangeOwnerDelta sheetTableDelta = StructuralPatch.RangeOwnerDelta.sheetTable("sheet-1", "sheet-table-1",
                new RangeRef("sheet-1", 4, 6, 0, 1), new RangeRef("sheet-1", 5, 8, 0, 1));
        return new StructuralPatch(StructuralPatch.VERSION, "rows.inserted", List.of(), List.of(), List.of(tableDelta, sheetTableDelta));
    }

    private List<RangeRef> preV5RangeImpacts(StructuralPatch patch, MutationDescriptorRegistry registry) {
        return registry.structuralImpactRanges(new StructuralPatch(StructuralPatch.VERSION, patch.mutationId(),
                patch.formulaOwnerDeltas(), patch.definedNameOwnerDeltas(), patch.rangeOwnerDeltas().stream()
                        .filter(delta -> !"sheet-table".equals(delta.ownerKind())).toList()));
    }
}
