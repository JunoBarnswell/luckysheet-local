package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.util.HashSet;
import java.util.List;
import java.util.Set;

/** Server-derived, versioned reference-owner delta attached to a committed structural intent. */
public record StructuralPatch(
        @JsonProperty("version") int version,
        @JsonProperty("mutationId") String mutationId,
        @JsonProperty("formulaOwnerDeltas") List<FormulaOwnerDelta> formulaOwnerDeltas
) {
    public static final int VERSION = 1;

    private record FormulaOwnerKey(String kind, String sheetId, Integer row, Integer column,
            String ruleKind, String ruleId, String field) { }

    @JsonCreator
    public StructuralPatch {
        if (version != VERSION) throw new IllegalArgumentException("Unsupported StructuralPatch version");
        if (mutationId == null || mutationId.isBlank()) throw new IllegalArgumentException("StructuralPatch mutationId is required");
        if (formulaOwnerDeltas == null) throw new IllegalArgumentException("StructuralPatch formulaOwnerDeltas are required");
        formulaOwnerDeltas = List.copyOf(formulaOwnerDeltas);
        Set<FormulaOwnerKey> ownerKeys = new HashSet<>();
        for (FormulaOwnerDelta delta : formulaOwnerDeltas) {
            FormulaOwnerKey key = "formula-cell".equals(delta.kind())
                    ? new FormulaOwnerKey(delta.kind(), delta.afterAddress().sheetId(), delta.afterAddress().row(),
                            delta.afterAddress().column(), null, null, null)
                    : new FormulaOwnerKey(delta.kind(), delta.sheetId(), null, null,
                            delta.ruleKind(), delta.ruleId(), delta.field());
            if (!ownerKeys.add(key)) throw new IllegalArgumentException("StructuralPatch contains duplicate formula owner deltas");
        }
    }

    public StructuralPatch inverse(String inverseMutationId) {
        List<FormulaOwnerDelta> inverse = formulaOwnerDeltas.stream()
                .map(FormulaOwnerDelta::inverse)
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

    @JsonInclude(JsonInclude.Include.NON_NULL)
    public record FormulaOwnerDelta(
            @JsonProperty("kind") String kind,
            @JsonProperty("beforeAddress") CellAddress beforeAddress,
            @JsonProperty("afterAddress") CellAddress afterAddress,
            @JsonProperty("before") FormulaOwnerState before,
            @JsonProperty("after") FormulaOwnerState after,
            @JsonProperty("sheetId") String sheetId,
            @JsonProperty("ruleKind") String ruleKind,
            @JsonProperty("ruleId") String ruleId,
            @JsonProperty("field") String field,
            @JsonProperty("beforeFormula") String beforeFormula,
            @JsonProperty("afterFormula") String afterFormula,
            @JsonProperty("beforeRanges") List<RangeRef> beforeRanges,
            @JsonProperty("afterRanges") List<RangeRef> afterRanges
    ) {
        @JsonCreator
        public FormulaOwnerDelta {
            if ("formula-cell".equals(kind)) {
                if (beforeAddress == null || afterAddress == null || before == null || after == null
                        || sheetId != null || ruleKind != null || ruleId != null || field != null
                        || beforeFormula != null || afterFormula != null || beforeRanges != null || afterRanges != null) {
                    throw new IllegalArgumentException("StructuralPatch formula-cell owner delta is incomplete or mixed with rule state");
                }
            } else if ("formula-rule".equals(kind)) {
                if (beforeAddress != null || afterAddress != null || before != null || after != null
                        || sheetId == null || sheetId.isBlank() || ruleId == null || ruleId.isBlank()
                        || ruleKind == null || !List.of("conditional-format", "data-validation").contains(ruleKind)
                        || field == null || !List.of("value1", "value2", "formula1", "formula2", "listSource.formula").contains(field)
                        || "conditional-format".equals(ruleKind) && !List.of("value1", "value2").contains(field)
                        || "data-validation".equals(ruleKind) && List.of("value1", "value2").contains(field)
                        || beforeFormula == null || afterFormula == null || beforeFormula.equals(afterFormula)
                        || beforeRanges == null || beforeRanges.isEmpty() || afterRanges == null || afterRanges.isEmpty()
                        || !validRanges(beforeRanges, sheetId) || !validRanges(afterRanges, sheetId)) {
                    throw new IllegalArgumentException("StructuralPatch formula-rule owner delta is incomplete or invalid");
                }
                beforeRanges = List.copyOf(beforeRanges);
                afterRanges = List.copyOf(afterRanges);
            } else {
                throw new IllegalArgumentException("Unsupported StructuralPatch formula owner kind");
            }
        }

        private static boolean validRanges(List<RangeRef> ranges, String sheetId) {
            return ranges.stream().allMatch(range -> range != null && sheetId.equals(range.sheetId())
                    && range.endRow() <= 1_048_575 && range.endColumn() <= 16_383);
        }

        public FormulaOwnerDelta(String kind, CellAddress beforeAddress, CellAddress afterAddress,
                FormulaOwnerState before, FormulaOwnerState after) {
            this(kind, beforeAddress, afterAddress, before, after, null, null, null, null, null, null, null, null);
        }

        public static FormulaOwnerDelta formulaRule(String sheetId, String ruleKind, String ruleId, String field,
                String beforeFormula, String afterFormula, List<RangeRef> beforeRanges, List<RangeRef> afterRanges) {
            return new FormulaOwnerDelta("formula-rule", null, null, null, null, sheetId, ruleKind, ruleId,
                    field, beforeFormula, afterFormula, beforeRanges, afterRanges);
        }

        public FormulaOwnerDelta inverse() {
            if ("formula-cell".equals(kind)) {
                return new FormulaOwnerDelta(kind, afterAddress, beforeAddress, after, before);
            }
            return formulaRule(sheetId, ruleKind, ruleId, field, afterFormula, beforeFormula, afterRanges, beforeRanges);
        }
    }
}
