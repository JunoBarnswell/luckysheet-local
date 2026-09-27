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

    @Test
    void upgradesV5HistoryWithoutRangeOnlyFormulaRuleFacts() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        RangeRef before = new RangeRef("sheet-1", 1, 3, 0, 2);
        RangeRef after = new RangeRef("sheet-1", 2, 4, 0, 2);
        StructuralPatch currentPatch = new StructuralPatch(StructuralPatch.VERSION, "rows.inserted",
                List.of(StructuralPatch.FormulaOwnerDelta.formulaRule("sheet-1", "conditional-format",
                        "rule-1", "value1", "=A1", "=A1", List.of(before), List.of(after))),
                List.of(), List.of());
        ObjectNode envelope = rawEnvelope(5, currentPatch, List.of());
        ((ObjectNode) envelope.path("mutations").get(0).path("structuralPatch"))
                .set("formulaOwnerDeltas", mapper.createArrayNode());

        new V4__TestMigration().rewriteMutationPatches(envelope, operation(),
                List.of(Optional.of(currentPatch)), registry, "unit-1", 1);

        assertEquals(mapper.valueToTree(currentPatch), envelope.path("mutations").get(0).get("structuralPatch"));
        assertEquals(mapper.valueToTree(registry.structuralImpactRanges(currentPatch)),
                envelope.path("mutations").get(0).get("structuralImpactRanges"));
    }

    @Test
    void upgradesV6AndV7HistoryWithFormulaAnchorOwnerFacts() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        RangeRef before = new RangeRef("sheet-1", 0, 3, 2, 2);
        RangeRef after = new RangeRef("sheet-1", 0, 5, 2, 2);
        StructuralPatch currentPatch = new StructuralPatch(StructuralPatch.VERSION, "rows.inserted",
                List.of(StructuralPatch.FormulaOwnerDelta.formulaRuleAnchor("sheet-1", "data-validation", "validation-1",
                        new StructuralPatch.CellAddress("sheet-1", 0, 0), new StructuralPatch.CellAddress("sheet-1", 1, 0))),
                List.of(), List.of(StructuralPatch.RangeOwnerDelta.validationListSource(
                        "sheet-1", "validation-1", before, after,
                        List.of(new RangeRef("sheet-1", 0, 0, 0, 0)), List.of(new RangeRef("sheet-1", 1, 1, 0, 0))),
                        StructuralPatch.RangeOwnerDelta.ruleRanges("conditional-format", "sheet-1", "cf-color-scale",
                                List.of(new RangeRef("sheet-1", 4, 4, 0, 2)), List.of(new RangeRef("sheet-1", 5, 5, 0, 2)))));

        ObjectNode v6Envelope = rawEnvelope(6, currentPatch, preV7RangeImpacts(currentPatch, registry));
        new V4__TestMigration().rewriteMutationPatches(v6Envelope, operation(),
                List.of(Optional.of(currentPatch)), registry, "unit-1", 1);
        assertEquals(mapper.valueToTree(currentPatch), v6Envelope.path("mutations").get(0).get("structuralPatch"));
        assertEquals(mapper.valueToTree(registry.structuralImpactRanges(currentPatch)),
                v6Envelope.path("mutations").get(0).get("structuralImpactRanges"));

        StructuralPatch preV8 = new StructuralPatch(StructuralPatch.VERSION, currentPatch.mutationId(),
                currentPatch.formulaOwnerDeltas().stream()
                        .filter(delta -> !"formula-rule-anchor".equals(delta.kind())).toList(),
                currentPatch.definedNameOwnerDeltas(), currentPatch.rangeOwnerDeltas().stream()
                        .filter(delta -> !List.of("conditional-format", "data-validation").contains(delta.ownerKind())).toList());
        ObjectNode v7Envelope = rawEnvelope(7, currentPatch, registry.structuralImpactRanges(preV8));
        new V4__TestMigration().rewriteMutationPatches(v7Envelope, operation(),
                List.of(Optional.of(currentPatch)), registry, "unit-1", 1);
        assertEquals(mapper.valueToTree(currentPatch), v7Envelope.path("mutations").get(0).get("structuralPatch"));

        ObjectNode v8Envelope = rawEnvelope(8, currentPatch, registry.structuralImpactRanges(currentPatch));
        new V4__TestMigration().rewriteMutationPatches(v8Envelope, operation(),
                List.of(Optional.of(currentPatch)), registry, "unit-1", 1);
        assertEquals(mapper.valueToTree(currentPatch), v8Envelope.path("mutations").get(0).get("structuralPatch"));

        StructuralPatch implicitAnchorPatch = new StructuralPatch(StructuralPatch.VERSION, "rows.inserted",
                List.of(StructuralPatch.FormulaOwnerDelta.formulaRuleAnchor("sheet-1", "data-validation", "validation-1",
                        null, new StructuralPatch.CellAddress("sheet-1", 1, 0))), List.of(), List.of());
        StructuralPatch preV9 = new StructuralPatch(StructuralPatch.VERSION, implicitAnchorPatch.mutationId(),
                List.of(), List.of(), List.of());
        ObjectNode v8ImplicitAnchorEnvelope = rawEnvelope(8, implicitAnchorPatch, registry.structuralImpactRanges(preV9));
        new V4__TestMigration().rewriteMutationPatches(v8ImplicitAnchorEnvelope, operation(),
                List.of(Optional.of(implicitAnchorPatch)), registry, "unit-1", 1);
        assertEquals(mapper.valueToTree(implicitAnchorPatch),
                v8ImplicitAnchorEnvelope.path("mutations").get(0).get("structuralPatch"));

        ObjectNode v9Envelope = rawEnvelope(9, implicitAnchorPatch, registry.structuralImpactRanges(implicitAnchorPatch));
        new V4__TestMigration().rewriteMutationPatches(v9Envelope, operation(),
                List.of(Optional.of(implicitAnchorPatch)), registry, "unit-1", 1);
        assertEquals(mapper.valueToTree(implicitAnchorPatch), v9Envelope.path("mutations").get(0).get("structuralPatch"));
    }

    @Test
    void backfillsLegacyWorksheetRenameOwnerPatches() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        StructuralPatch renamePatch = new StructuralPatch(StructuralPatch.VERSION,
                "sheet.rename", List.of(), List.of(), List.of());
        ObjectNode envelope = mapper.createObjectNode();
        ObjectNode mutationNode = envelope.putArray("mutations").addObject()
                .put("id", "sheet.rename").put("sheetId", "sheet-1");
        mutationNode.set("structuralImpactRanges", mapper.createArrayNode());
        Instant now = Instant.parse("2026-09-26T00:00:00Z");
        CommittedOperationEnvelope operation = new CommittedOperationEnvelope(
                "session-1", OperationEnvelope.SCHEMA, "operation-1", "unit-1", "actor-1",
                OperationOrigin.CLIENT, 1, 0, 1,
                List.of(new CommittedOperationMutation("sheet.rename", "sheet-1", mapper.createObjectNode(), List.of())), now, now);

        new V4__TestMigration().rewriteMutationPatches(envelope, operation,
                List.of(Optional.of(renamePatch)), registry, "unit-1", 1);

        assertEquals(mapper.valueToTree(renamePatch), envelope.path("mutations").get(0).get("structuralPatch"));
        assertEquals(mapper.createArrayNode(), envelope.path("mutations").get(0).get("structuralImpactRanges"));
    }

    private ObjectNode rawEnvelope(int version, StructuralPatch currentPatch, List<RangeRef> impact) {
        ObjectNode legacyPatch = (ObjectNode) mapper.valueToTree(currentPatch);
        legacyPatch.put("version", version);
        if (version < 9) {
            ArrayNode priorFormulaOwners = mapper.createArrayNode();
            for (JsonNode delta : legacyPatch.path("formulaOwnerDeltas")) {
                if (!"formula-rule-anchor".equals(delta.path("kind").asText())
                        || delta.hasNonNull("beforeAddress") && delta.hasNonNull("afterAddress")) priorFormulaOwners.add(delta.deepCopy());
            }
            legacyPatch.set("formulaOwnerDeltas", priorFormulaOwners);
        }
        if (version < 8) {
            ArrayNode priorFormulaOwners = mapper.createArrayNode();
            for (JsonNode delta : legacyPatch.path("formulaOwnerDeltas")) {
                if (!"formula-rule-anchor".equals(delta.path("kind").asText())) priorFormulaOwners.add(delta.deepCopy());
            }
            legacyPatch.set("formulaOwnerDeltas", priorFormulaOwners);
            ArrayNode priorRangeOwners = mapper.createArrayNode();
            for (JsonNode delta : legacyPatch.path("rangeOwnerDeltas")) {
                if (!List.of("conditional-format", "data-validation").contains(delta.path("ownerKind").asText())) {
                    priorRangeOwners.add(delta.deepCopy());
                }
            }
            legacyPatch.set("rangeOwnerDeltas", priorRangeOwners);
        }
        if (version < 7) {
            ArrayNode priorRangeOwners = mapper.createArrayNode();
            for (JsonNode delta : legacyPatch.path("rangeOwnerDeltas")) {
                if (!"validation-list-source".equals(delta.path("ownerKind").asText())) priorRangeOwners.add(delta.deepCopy());
            }
            legacyPatch.set("rangeOwnerDeltas", priorRangeOwners);
        }
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

    private List<RangeRef> preV7RangeImpacts(StructuralPatch patch, MutationDescriptorRegistry registry) {
        return registry.structuralImpactRanges(new StructuralPatch(StructuralPatch.VERSION, patch.mutationId(),
                patch.formulaOwnerDeltas().stream().filter(delta -> !"formula-rule-anchor".equals(delta.kind())).toList(),
                patch.definedNameOwnerDeltas(), patch.rangeOwnerDeltas().stream()
                        .filter(delta -> !List.of("validation-list-source", "conditional-format", "data-validation")
                                .contains(delta.ownerKind())).toList()));
    }
}
