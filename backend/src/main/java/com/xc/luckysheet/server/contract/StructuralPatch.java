package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;

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
        @JsonProperty("definedNameOwnerDeltas") List<DefinedNameOwnerDelta> definedNameOwnerDeltas,
        @JsonProperty("rangeOwnerDeltas") List<RangeOwnerDelta> rangeOwnerDeltas
) {
    public static final int VERSION = 4;

    public record FormulaOwnerKey(String kind, String sheetId, Integer row, Integer column,
            String ruleKind, String ruleId, String field, String ownerKind, String ownerId,
            String fieldId, String viewId, String templateId) { }
    public record DefinedNameOwnerKey(String scope, String normalizedName, String sheetId) { }
    public record RangeOwnerKey(String ownerKind, String sheetId, String ownerId, String regionId) { }

    @JsonCreator
    public StructuralPatch {
        if (version != VERSION) throw new IllegalArgumentException("Unsupported StructuralPatch version");
        if (mutationId == null || mutationId.isBlank()) throw new IllegalArgumentException("StructuralPatch mutationId is required");
        if (formulaOwnerDeltas == null) throw new IllegalArgumentException("StructuralPatch formulaOwnerDeltas are required");
        if (definedNameOwnerDeltas == null) throw new IllegalArgumentException("StructuralPatch definedNameOwnerDeltas are required");
        if (rangeOwnerDeltas == null) throw new IllegalArgumentException("StructuralPatch rangeOwnerDeltas are required");
        formulaOwnerDeltas = List.copyOf(formulaOwnerDeltas);
        definedNameOwnerDeltas = List.copyOf(definedNameOwnerDeltas);
        rangeOwnerDeltas = List.copyOf(rangeOwnerDeltas);
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
        Set<RangeOwnerKey> rangeOwnerKeys = new HashSet<>();
        for (RangeOwnerDelta delta : rangeOwnerDeltas) {
            RangeOwnerKey key = rangeOwnerKey(delta);
            if (!rangeOwnerKeys.add(key)) throw new IllegalArgumentException("StructuralPatch contains duplicate range-owner deltas");
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

    public StructuralPatch inverse(String inverseMutationId) {
        List<FormulaOwnerDelta> inverse = formulaOwnerDeltas.stream()
                .map(FormulaOwnerDelta::inverse)
                .toList();
        List<DefinedNameOwnerDelta> inverseDefinedNames = definedNameOwnerDeltas.stream()
                .map(DefinedNameOwnerDelta::inverse)
                .toList();
        List<RangeOwnerDelta> inverseRanges = rangeOwnerDeltas.stream()
                .map(RangeOwnerDelta::inverse)
                .toList();
        return new StructuralPatch(VERSION, inverseMutationId, inverse, inverseDefinedNames, inverseRanges);
    }

    public static RangeOwnerKey rangeOwnerKey(RangeOwnerDelta delta) {
        if (delta == null) throw new IllegalArgumentException("StructuralPatch range-owner delta is required");
        return new RangeOwnerKey(delta.ownerKind(), delta.sheetId(), delta.ownerId(), delta.regionId());
    }

    @JsonInclude(JsonInclude.Include.NON_NULL)
    public record RangeOwnerDelta(
            @JsonProperty("ownerKind") String ownerKind,
            @JsonProperty("sheetId") String sheetId,
            @JsonProperty("regionId") String regionId,
            @JsonProperty("ownerId") String ownerId,
            @JsonProperty("before") JsonNode before,
            @JsonProperty("after") JsonNode after
    ) {
        @JsonCreator
        public RangeOwnerDelta {
            if (before == null || after == null) throw new IllegalArgumentException("StructuralPatch range-owner states are required");
            before = before.deepCopy();
            after = after.deepCopy();
            if ("data-region".equals(ownerKind)) {
                if (sheetId == null || sheetId.isBlank() || regionId == null || regionId.isBlank() || ownerId != null
                        || !validRegionState(before, sheetId) || !validRegionState(after, sheetId)
                        || !sameExtent(regionRange(before), regionRange(after))
                        || before.equals(after)) {
                    throw new IllegalArgumentException("StructuralPatch data-region range-owner delta is invalid");
                }
            } else if (List.of("workbook-table", "data-source").contains(ownerKind)) {
                if (sheetId != null || regionId != null || ownerId == null || ownerId.isBlank()
                        || !validRange(before) || !validRange(after)
                        || !range(before).sheetId().equals(range(after).sheetId())
                        || !sameExtent(range(before), range(after)) || before.equals(after)) {
                    throw new IllegalArgumentException("StructuralPatch range-owner delta is invalid");
                }
            } else {
                throw new IllegalArgumentException("Unsupported StructuralPatch range-owner kind");
            }
        }

        public static RangeOwnerDelta dataRegion(String sheetId, String regionId,
                RangeRef beforeRange, int beforeHeaderRow, RangeRef afterRange, int afterHeaderRow) {
            ObjectNode before = JsonNodeFactory.instance.objectNode();
            before.set("range", rangeNode(beforeRange));
            before.put("headerRow", beforeHeaderRow);
            ObjectNode after = JsonNodeFactory.instance.objectNode();
            after.set("range", rangeNode(afterRange));
            after.put("headerRow", afterHeaderRow);
            return new RangeOwnerDelta("data-region", sheetId, regionId, null, before, after);
        }

        public static RangeOwnerDelta range(String ownerKind, String ownerId, RangeRef before, RangeRef after) {
            if (!List.of("workbook-table", "data-source").contains(ownerKind)) {
                throw new IllegalArgumentException("Unsupported StructuralPatch range owner kind");
            }
            return new RangeOwnerDelta(ownerKind, null, null, ownerId, rangeNode(before), rangeNode(after));
        }

        public RangeRef beforeRange() { return "data-region".equals(ownerKind) ? regionRange(before) : range(before); }
        public RangeRef afterRange() { return "data-region".equals(ownerKind) ? regionRange(after) : range(after); }
        public int beforeHeaderRow() { return regionHeaderRow(before); }
        public int afterHeaderRow() { return regionHeaderRow(after); }

        public RangeOwnerDelta inverse() {
            return new RangeOwnerDelta(ownerKind, sheetId, regionId, ownerId, after, before);
        }

        private static boolean validRegionState(JsonNode state, String sheetId) {
            if (!state.isObject() || !hasExactFields(state, Set.of("range", "headerRow"))
                    || !validRange(state.get("range")) || !sheetId.equals(range(state.get("range")).sheetId())
                    || !state.get("headerRow").isIntegralNumber() || !state.get("headerRow").canConvertToInt()) return false;
            int headerRow = state.get("headerRow").intValue();
            RangeRef range = range(state.get("range"));
            return headerRow >= range.startRow() && headerRow <= range.endRow();
        }

        private static boolean validRange(JsonNode value) {
            if (value == null || !value.isObject()
                    || !hasExactFields(value, Set.of("sheetId", "startRow", "endRow", "startColumn", "endColumn"))) return false;
            JsonNode sheetId = value.get("sheetId");
            JsonNode startRow = value.get("startRow");
            JsonNode endRow = value.get("endRow");
            JsonNode startColumn = value.get("startColumn");
            JsonNode endColumn = value.get("endColumn");
            return sheetId != null && sheetId.isTextual() && !sheetId.asText().isBlank()
                    && startRow != null && startRow.isIntegralNumber() && startRow.canConvertToInt()
                    && endRow != null && endRow.isIntegralNumber() && endRow.canConvertToInt()
                    && startColumn != null && startColumn.isIntegralNumber() && startColumn.canConvertToInt()
                    && endColumn != null && endColumn.isIntegralNumber() && endColumn.canConvertToInt()
                    && startRow.intValue() >= 0 && endRow.intValue() >= startRow.intValue() && endRow.intValue() <= 1_048_575
                    && startColumn.intValue() >= 0 && endColumn.intValue() >= startColumn.intValue() && endColumn.intValue() <= 16_383;
        }

        private static boolean hasExactFields(JsonNode node, Set<String> expected) {
            Set<String> actual = new HashSet<>();
            node.fieldNames().forEachRemaining(actual::add);
            return actual.equals(expected);
        }

        private static boolean sameExtent(RangeRef left, RangeRef right) {
            return left.endRow() - left.startRow() == right.endRow() - right.startRow()
                    && left.endColumn() - left.startColumn() == right.endColumn() - right.startColumn();
        }

        private static RangeRef range(JsonNode node) {
            return new RangeRef(node.path("sheetId").asText(), node.path("startRow").intValue(), node.path("endRow").intValue(),
                    node.path("startColumn").intValue(), node.path("endColumn").intValue());
        }

        private static RangeRef regionRange(JsonNode state) { return range(state.path("range")); }
        private static int regionHeaderRow(JsonNode state) { return state.path("headerRow").intValue(); }

        private static ObjectNode rangeNode(RangeRef range) {
            if (range == null || range.endRow() > 1_048_575 || range.endColumn() > 16_383) {
                throw new IllegalArgumentException("StructuralPatch range-owner range is outside worksheet bounds");
            }
            ObjectNode node = JsonNodeFactory.instance.objectNode();
            node.put("sheetId", range.sheetId());
            node.put("startRow", range.startRow());
            node.put("endRow", range.endRow());
            node.put("startColumn", range.startColumn());
            node.put("endColumn", range.endColumn());
            return node;
        }
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
