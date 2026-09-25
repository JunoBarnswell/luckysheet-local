package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.util.List;

/** Server-derived, versioned reference-owner delta attached to a committed structural intent. */
public record StructuralPatch(
        @JsonProperty("version") int version,
        @JsonProperty("mutationId") String mutationId,
        @JsonProperty("formulaOwnerDeltas") List<FormulaOwnerDelta> formulaOwnerDeltas
) {
    public static final int VERSION = 1;

    @JsonCreator
    public StructuralPatch {
        if (version != VERSION) throw new IllegalArgumentException("Unsupported StructuralPatch version");
        if (mutationId == null || mutationId.isBlank()) throw new IllegalArgumentException("StructuralPatch mutationId is required");
        if (formulaOwnerDeltas == null) throw new IllegalArgumentException("StructuralPatch formulaOwnerDeltas are required");
        formulaOwnerDeltas = List.copyOf(formulaOwnerDeltas);
    }

    public StructuralPatch inverse(String inverseMutationId) {
        List<FormulaOwnerDelta> inverse = formulaOwnerDeltas.stream()
                .map(delta -> new FormulaOwnerDelta(
                        delta.kind(),
                        delta.afterAddress(),
                        delta.beforeAddress(),
                        delta.after(),
                        delta.before()))
                .toList();
        return new StructuralPatch(VERSION, inverseMutationId, inverse);
    }

    public record CellAddress(
            @JsonProperty("sheetId") String sheetId,
            @JsonProperty("row") int row,
            @JsonProperty("column") int column
    ) {
        @JsonCreator
        public CellAddress {
            if (sheetId == null || sheetId.isBlank()) throw new IllegalArgumentException("StructuralPatch owner sheetId is required");
            if (row < 0 || row > 1_048_575 || column < 0 || column > 16_383) {
                throw new IllegalArgumentException("StructuralPatch owner coordinates must be within worksheet bounds");
            }
        }
    }

    @JsonInclude(JsonInclude.Include.ALWAYS)
    public record FormulaOwnerState(
            @JsonProperty("formula") String formula,
            @JsonProperty("sourceFormula") String sourceFormula,
            @JsonProperty("barcodeFormula") String barcodeFormula
    ) { }

    public record FormulaOwnerDelta(
            @JsonProperty("kind") String kind,
            @JsonProperty("beforeAddress") CellAddress beforeAddress,
            @JsonProperty("afterAddress") CellAddress afterAddress,
            @JsonProperty("before") FormulaOwnerState before,
            @JsonProperty("after") FormulaOwnerState after
    ) {
        @JsonCreator
        public FormulaOwnerDelta {
            if (!"formula-cell".equals(kind)) throw new IllegalArgumentException("Unsupported StructuralPatch formula owner kind");
            if (beforeAddress == null || afterAddress == null || before == null || after == null) {
                throw new IllegalArgumentException("StructuralPatch formula owner delta is incomplete");
            }
        }
    }
}
