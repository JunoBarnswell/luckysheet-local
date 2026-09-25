package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;
import java.util.List;

@JsonInclude(JsonInclude.Include.NON_NULL)
public record CommittedOperationMutation(
        @JsonProperty("id") String id,
        @JsonProperty("sheetId") String sheetId,
        @JsonProperty("params") JsonNode params,
        @JsonProperty("affectedRanges") List<RangeRef> affectedRanges,
        @JsonProperty("structuralImpactRanges") List<RangeRef> structuralImpactRanges,
        @JsonProperty("structuralPatch") StructuralPatch structuralPatch
) {
    @JsonCreator
    public CommittedOperationMutation {
        if (id == null || id.isBlank()) throw new IllegalArgumentException("mutation id is required");
        if (sheetId == null || sheetId.isBlank()) throw new IllegalArgumentException("mutation sheetId is required");
        if (params == null) throw new IllegalArgumentException("mutation params are required");
        if (affectedRanges == null) throw new IllegalArgumentException("affectedRanges are server-owned");
        affectedRanges = List.copyOf(affectedRanges);
        structuralImpactRanges = structuralImpactRanges == null ? List.of() : List.copyOf(structuralImpactRanges);
    }

    public CommittedOperationMutation(String id, String sheetId, JsonNode params, List<RangeRef> affectedRanges) {
        this(id, sheetId, params, affectedRanges, List.of(), null);
    }

    public CommittedOperationMutation(String id, String sheetId, JsonNode params, List<RangeRef> affectedRanges, StructuralPatch structuralPatch) {
        this(id, sheetId, params, affectedRanges, List.of(), structuralPatch);
    }

    public static CommittedOperationMutation from(OperationMutation mutation, List<RangeRef> ranges) {
        return from(mutation, ranges, List.of(), null);
    }

    public static CommittedOperationMutation from(OperationMutation mutation, List<RangeRef> ranges, StructuralPatch structuralPatch) {
        return from(mutation, ranges, List.of(), structuralPatch);
    }

    public static CommittedOperationMutation from(
            OperationMutation mutation,
            List<RangeRef> ranges,
            List<RangeRef> structuralImpactRanges,
            StructuralPatch structuralPatch
    ) {
        return new CommittedOperationMutation(mutation.id(), mutation.sheetId(), mutation.params(), ranges, structuralImpactRanges, structuralPatch);
    }
}
