package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;

/** Server-derived, versioned reference-owner delta attached to a committed structural intent. */
@JsonInclude(JsonInclude.Include.ALWAYS)
public record StructuralPatch(
        @JsonProperty("version") int version,
        @JsonProperty("mutationId") String mutationId,
        @JsonProperty("formulaOwnerDeltas") List<FormulaOwnerDelta> formulaOwnerDeltas,
        @JsonProperty("definedNameOwnerDeltas") List<DefinedNameOwnerDelta> definedNameOwnerDeltas
) {
    public static final int VERSION = 3;

    public record FormulaOwnerKey(String kind, String sheetId, Integer row, Integer column,
            String ruleKind, String ruleId, String field, String ownerKind, String ownerId,
            String fieldId, String viewId, String templateId) { }
    public record DefinedNameOwnerKey(String scope, String normalizedName, String sheetId) { }

    @JsonCreator
    public StructuralPatch {
        if (version != VERSION) throw new IllegalArgumentException("Unsupported StructuralPatch version");
        if (mutationId == null || mutationId.isBlank()) throw new IllegalArgumentException("StructuralPatch mutationId is required");
        if (formulaOwnerDeltas == null) throw new IllegalArgumentException("StructuralPatch formulaOwnerDeltas are required");
        if (definedNameOwnerDeltas == null) throw new IllegalArgumentException("StructuralPatch definedNameOwnerDeltas are required");
        formulaOwnerDeltas = List.copyOf(formulaOwnerDeltas);
        definedNameOwnerDeltas = List.copyOf(definedNameOwnerDeltas);
        Set<FormulaOwnerKey> ownerKeys = new HashSet<>();
        for (FormulaOwnerDelta delta : formulaOwnerDeltas) {
            FormulaOwnerKey key = formulaOwnerKey(delta);
            if (!ownerKeys.add(key)) throw new IllegalArgumentException("StructuralPatch contains duplicate formula owner deltas");
        }
        Set<DefinedNameOwnerKey> definedNameKeys = new HashSet<>();
        for (DefinedNameOwnerDelta delta : definedNameOwnerDeltas) {
            DefinedNameOwnerKey key = definedNameOwnerKey(delta);
            if (!definedNameKeys.add(key)) throw new IllegalArgumentException("StructuralPatch contains duplicate defined-name owner deltas");
        }
    }

    public static FormulaOwnerKey formulaOwnerKey(FormulaOwnerDelta delta) {
        if (delta == null) throw new IllegalArgumentException("StructuralPatch formula owner delta is required");
        return switch (delta.kind()) {
            case "formula-cell" -> {
                CellAddress address = delta.afterAddress();
                yield new FormulaOwnerKey(delta.kind(), address.sheetId(), address.row(), address.column(),
                        null, null, null, null, null, null, null, null);
            }
            case "formula-rule" -> new FormulaOwnerKey(delta.kind(), delta.sheetId(), null, null,
                    delta.ruleKind(), delta.ruleId(), delta.field(), null, null, null, null, null);
            case "formula-object" -> new FormulaOwnerKey(delta.kind(), delta.sheetId(), null, null,
                    null, null, delta.field(), delta.ownerKind(), delta.ownerId(),
                    delta.fieldId(), delta.viewId(), delta.templateId());
            default -> throw new IllegalArgumentException("Unsupported StructuralPatch formula owner kind");
        };
    }

    public static DefinedNameOwnerKey definedNameOwnerKey(DefinedNameOwnerDelta delta) {
        if (delta == null) throw new IllegalArgumentException("StructuralPatch defined-name owner delta is required");
        DefinedNameOwnerIdentity owner = delta.owner();
        return new DefinedNameOwnerKey(owner.scope(), owner.name().toUpperCase(Locale.ROOT), owner.sheetId());
    }

    public StructuralPatch(int version, String mutationId, List<FormulaOwnerDelta> formulaOwnerDeltas) {
        this(version, mutationId, formulaOwnerDeltas, List.of());
    }

    public StructuralPatch inverse(String inverseMutationId) {
        List<FormulaOwnerDelta> inverse = formulaOwnerDeltas.stream()
                .map(FormulaOwnerDelta::inverse)
                .toList();
        List<DefinedNameOwnerDelta> inverseDefinedNames = definedNameOwnerDeltas.stream()
                .map(DefinedNameOwnerDelta::inverse)
                .toList();
        return new StructuralPatch(VERSION, inverseMutationId, inverse, inverseDefinedNames);
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

    @JsonInclude(JsonInclude.Include.NON_NULL)
    public record DefinedNameOwnerIdentity(
            @JsonProperty("scope") String scope,
            @JsonProperty("name") String name,
            @JsonProperty("sheetId") String sheetId
    ) {
        @JsonCreator
        public DefinedNameOwnerIdentity {
            if (!List.of("workbook", "sheet").contains(scope)
                    || name == null || !isEcmaScriptTrimmed(name) || name.length() > 255
                    || !name.matches("^[A-Za-z_\\\\][A-Za-z0-9_.]*$")
                    || ("sheet".equals(scope) && (sheetId == null || !isEcmaScriptTrimmed(sheetId)))
                    || ("workbook".equals(scope) && sheetId != null)) {
                throw new IllegalArgumentException("StructuralPatch defined-name owner identity is invalid");
            }
        }
    }

    @JsonInclude(JsonInclude.Include.NON_NULL)
    public record DefinedNameState(
            @JsonProperty("name") String name,
            @JsonProperty("formula") String formula,
            @JsonProperty("scope") String scope,
            @JsonProperty("sheetId") String sheetId,
            @JsonProperty("anchor") CellAddress anchor
    ) {
        @JsonCreator
        public DefinedNameState {
            new DefinedNameOwnerIdentity(scope, name, sheetId);
            if (formula == null || !isEcmaScriptTrimmed(formula) || formula.length() > 32_767) {
                throw new IllegalArgumentException("StructuralPatch defined-name formula is invalid");
            }
        }
    }

    public record DefinedNameOwnerDelta(
            @JsonProperty("owner") DefinedNameOwnerIdentity owner,
            @JsonProperty("before") DefinedNameState before,
            @JsonProperty("after") DefinedNameState after
    ) {
        @JsonCreator
        public DefinedNameOwnerDelta {
            if (owner == null || before == null || after == null
                    || !owner.equals(identity(before)) || !owner.equals(identity(after))
                    || before.equals(after)) {
                throw new IllegalArgumentException("StructuralPatch defined-name owner delta is incomplete or invalid");
            }
        }

        public DefinedNameOwnerDelta inverse() {
            return new DefinedNameOwnerDelta(owner, after, before);
        }

        private static DefinedNameOwnerIdentity identity(DefinedNameState state) {
            return new DefinedNameOwnerIdentity(state.scope(), state.name(), state.sheetId());
        }
    }

    private static boolean isEcmaScriptTrimmed(String value) {
        if (value.isEmpty()) return false;
        return !isEcmaScriptTrimWhitespace(value.codePointAt(0))
                && !isEcmaScriptTrimWhitespace(value.codePointBefore(value.length()));
    }

    private static boolean isEcmaScriptTrimWhitespace(int codePoint) {
        return (codePoint >= 0x0009 && codePoint <= 0x000D)
                || codePoint == 0x0020 || codePoint == 0x00A0 || codePoint == 0x1680
                || (codePoint >= 0x2000 && codePoint <= 0x200A)
                || codePoint == 0x2028 || codePoint == 0x2029 || codePoint == 0x202F
                || codePoint == 0x205F || codePoint == 0x3000 || codePoint == 0xFEFF;
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
            @JsonProperty("afterRanges") List<RangeRef> afterRanges,
            @JsonProperty("ownerKind") String ownerKind,
            @JsonProperty("payloadId") String ownerId,
            @JsonProperty("fieldId") String fieldId,
            @JsonProperty("viewId") String viewId,
            @JsonProperty("templateId") String templateId
    ) {
        @JsonCreator
        public FormulaOwnerDelta {
            if ("formula-cell".equals(kind)) {
                if (beforeAddress == null || afterAddress == null || before == null || after == null
                        || sheetId != null || ruleKind != null || ruleId != null || field != null
                        || ownerKind != null || ownerId != null || fieldId != null || viewId != null || templateId != null
                        || beforeFormula != null || afterFormula != null || beforeRanges != null || afterRanges != null) {
                    throw new IllegalArgumentException("StructuralPatch formula-cell owner delta is incomplete or mixed with rule state");
                }
            } else if ("formula-rule".equals(kind)) {
                if (beforeAddress != null || afterAddress != null || before != null || after != null
                        || ownerKind != null || ownerId != null || fieldId != null || viewId != null || templateId != null
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
            } else if ("formula-object".equals(kind)) {
                if (beforeAddress != null || afterAddress != null || before != null || after != null
                        || ruleKind != null || ruleId != null || beforeRanges != null || afterRanges != null
                        || beforeFormula == null || afterFormula == null || beforeFormula.equals(afterFormula)
                        || !validFormulaObjectOwner(ownerKind, sheetId, ownerId, fieldId, viewId, templateId, field)) {
                    throw new IllegalArgumentException("StructuralPatch formula-object owner delta is incomplete or invalid");
                }
            } else {
                throw new IllegalArgumentException("Unsupported StructuralPatch formula owner kind");
            }
        }

        private static boolean validRanges(List<RangeRef> ranges, String sheetId) {
            return ranges.stream().allMatch(range -> range != null && sheetId.equals(range.sheetId())
                    && range.endRow() <= 1_048_575 && range.endColumn() <= 16_383);
        }

        private static boolean validFormulaObjectOwner(String ownerKind, String sheetId, String payloadId,
                String fieldId, String viewId, String templateId, String field) {
            return switch (ownerKind == null ? "" : ownerKind) {
                case "chart-text" -> sheetId != null && !sheetId.isBlank() && payloadId != null && !payloadId.isBlank()
                        && fieldId == null && viewId == null && templateId == null
                        && List.of("titleText.linkedFormula", "legend.text.linkedFormula",
                                "categoryAxis.titleText.linkedFormula", "valueAxis.titleText.linkedFormula",
                                "secondaryCategoryAxis.titleText.linkedFormula", "secondaryValueAxis.titleText.linkedFormula",
                                "dataTable.font.linkedFormula").contains(field);
                case "shape-property" -> sheetId != null && !sheetId.isBlank() && payloadId != null && !payloadId.isBlank()
                        && field == null && fieldId == null && viewId == null && templateId == null;
                case "table-sheet-column" -> sheetId != null && !sheetId.isBlank() && fieldId != null && !fieldId.isBlank()
                        && payloadId == null && field == null && viewId == null && templateId == null;
                case "data-view-field" -> viewId != null && !viewId.isBlank() && fieldId != null && !fieldId.isBlank()
                        && sheetId == null && payloadId == null && field == null && templateId == null;
                case "cell-style-template" -> templateId != null && !templateId.isBlank() && fieldId == null
                        && viewId == null && sheetId == null && payloadId == null
                        && List.of("formula1", "formula2", "listSource.formula").contains(field);
                default -> false;
            };
        }

        public FormulaOwnerDelta(String kind, CellAddress beforeAddress, CellAddress afterAddress,
                FormulaOwnerState before, FormulaOwnerState after) {
            this(kind, beforeAddress, afterAddress, before, after, null, null, null, null, null,
                    null, null, null, null, null, null, null, null);
        }

        public static FormulaOwnerDelta formulaRule(String sheetId, String ruleKind, String ruleId, String field,
                String beforeFormula, String afterFormula, List<RangeRef> beforeRanges, List<RangeRef> afterRanges) {
            return new FormulaOwnerDelta("formula-rule", null, null, null, null, sheetId, ruleKind, ruleId,
                    field, beforeFormula, afterFormula, beforeRanges, afterRanges, null, null, null, null, null);
        }

        public static FormulaOwnerDelta formulaObject(String sheetId, String ownerKind, String ownerId, String field,
                String beforeFormula, String afterFormula) {
            return new FormulaOwnerDelta("formula-object", null, null, null, null, sheetId, null, null,
                    field, beforeFormula, afterFormula, null, null, ownerKind, ownerId, null, null, null);
        }

        public static FormulaOwnerDelta formulaObject(String ownerKind, String sheetId, String payloadId,
                String fieldId, String viewId, String templateId, String field, String beforeFormula, String afterFormula) {
            return new FormulaOwnerDelta("formula-object", null, null, null, null, sheetId, null, null,
                    field, beforeFormula, afterFormula, null, null, ownerKind, payloadId, fieldId, viewId, templateId);
        }

        public FormulaOwnerDelta inverse() {
            if ("formula-cell".equals(kind)) {
                return new FormulaOwnerDelta(kind, afterAddress, beforeAddress, after, before);
            }
            if ("formula-rule".equals(kind)) return formulaRule(sheetId, ruleKind, ruleId, field, afterFormula, beforeFormula, afterRanges, beforeRanges);
            return new FormulaOwnerDelta("formula-object", null, null, null, null, sheetId, null, null,
                    field, afterFormula, beforeFormula, null, null, ownerKind, ownerId, fieldId, viewId, templateId);
        }
    }
}
